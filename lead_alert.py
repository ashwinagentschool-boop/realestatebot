"""
Hyderabad real estate — frequent lead alerts.

Designed to run every 2 hours. Sends Telegram messages for newly detected
end-user leads. No Excel file.

Pipeline per post:
    1. Fetch posts from configured subreddits (24h lookback)
    2. Skip post IDs we've already processed (state file dedup)
    3. Claude #1 -> is_lead?
    4. If lead:
         - Reddit -> fetch user history
         - Claude #2 -> end_user / agent / unclear
    5. Filter: include only end_user OR (unclear AND quality >= 8)
    6. Send Telegram messages (no Excel)
    7. Update state file with all newly processed post IDs

State file (lead_state.json) format:
    {
      "seen": {
        "<post_id>": <unix_timestamp_when_added>
      }
    }
    IDs older than 7 days are pruned each run to keep the file small.

Required environment variables:
    TELEGRAM_BOT_TOKEN
    TELEGRAM_CHAT_ID
    ANTHROPIC_API_KEY
"""

import json
import os
import re
import time
import traceback
from collections import Counter
from datetime import datetime, timezone, timedelta

import requests
from anthropic import Anthropic

# Optional: load .env when running locally / on Pi
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


# =================== CONFIG ===================

SUBREDDITS = ["hyderabadrealestate", "Hyderabad_highrises"]
LOOKBACK_HOURS = 24
POST_LIMIT = 100  # per subreddit
USER_AGENT = "reddit-telegram-leads/1.0"

CLAUDE_MODEL = "claude-opus-4-7"
CLAUDE_MAX_TOKENS = 500
CLAUDE_TIMEOUT_SEC = 30

USER_HISTORY_LIMIT = 25
STATE_FILE = "lead_state.json"
STATE_TTL_DAYS = 7  # prune IDs older than this

UNCLEAR_QUALITY_THRESHOLD = 8

BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]
TG_SEND_MSG = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"

IST = timezone(timedelta(hours=5, minutes=30))

anthropic_client = Anthropic()

RE_SUBREDDITS = {
    "hyderabadrealestate", "hyderabad_highrises", "realestate", "indiarealestate",
    "indianrealestate", "realestateindia", "bangalorerealestate",
    "mumbai_realestate", "delhirealestate", "chennairealestate",
    "puneproperty", "realestateinvesting", "propertyinvesting",
}

PROMO_PATTERNS = [
    r"\bdm\s+me\b", r"\bdm\s+for\b", r"\bpm\s+me\b",
    r"\bcontact\s+me\b", r"\bcall\s+me\b", r"\bcontact\s+for\s+details?\b",
    r"\bfeel\s+free\s+to\s+(?:dm|message|contact|reach)",
    r"\b100\s*%\s+genuine\b", r"\bgenuine\s+seller\b",
    r"\bbest\s+deal\b", r"\bpremium\s+location\b",
    r"\bready\s+to\s+(?:occupy|move)\b", r"\bspot\s+booking\b",
    r"\binvestors?\s+wanted\b", r"\bfor\s+booking\s+call\b",
    r"\bhandling\s+sales\b", r"\bavailable\s+for\s+sale\b",
    r"\bwhatsapp\s*[:\-]?\s*\+?\d", r"\b\+?91[\s\-]?\d{10}\b",
    r"\b[6-9]\d{9}\b",
]


# =================== STATE FILE ===================

def load_state():
    if not os.path.exists(STATE_FILE):
        return {"seen": {}}
    try:
        with open(STATE_FILE, "r") as f:
            return json.load(f)
    except Exception as e:
        print(f"[state] could not load {STATE_FILE}: {e}")
        return {"seen": {}}


def save_state(state):
    # Prune old IDs
    cutoff = time.time() - (STATE_TTL_DAYS * 86400)
    state["seen"] = {pid: ts for pid, ts in state["seen"].items() if ts >= cutoff}
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


# =================== REDDIT FETCH ===================

def reddit_get(url):
    headers = {"User-Agent": USER_AGENT}
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code in (404, 403):
        return None
    resp.raise_for_status()
    return resp.json()


def fetch_recent_posts():
    cutoff = time.time() - (LOOKBACK_HOURS * 3600)
    posts = []
    seen_ids = set()  # in-run dedup for cross-posts

    for sub in SUBREDDITS:
        url = f"https://www.reddit.com/r/{sub}/new.json?limit={POST_LIMIT}"
        try:
            data = reddit_get(url)
        except Exception as e:
            print(f"[reddit-fetch-fail] r/{sub}: {e}")
            continue
        if not data:
            print(f"[reddit-fetch] r/{sub} returned no data")
            continue

        for child in data["data"]["children"]:
            d = child["data"]
            if d.get("created_utc", 0) < cutoff:
                continue

            pid = d.get("id", "")
            if pid in seen_ids:
                continue
            seen_ids.add(pid)

            posts.append({
                "id": pid,
                "title": d.get("title", "") or "",
                "selftext": d.get("selftext", "") or "",
                "author": d.get("author", "unknown") or "unknown",
                "score": d.get("score", 0),
                "num_comments": d.get("num_comments", 0),
                "permalink": f"https://reddit.com{d.get('permalink', '')}",
                "url": d.get("url", "") or "",
                "is_self": d.get("is_self", True),
                "flair": d.get("link_flair_text", "") or "",
                "created_utc": d.get("created_utc", 0),
                "subreddit": d.get("subreddit", sub) or sub,
            })

    posts.sort(key=lambda x: x["created_utc"], reverse=True)
    return posts


def fetch_user_history(username):
    if username in ("[deleted]", "unknown", "AutoModerator"):
        return None

    about = reddit_get(f"https://www.reddit.com/user/{username}/about.json")
    if not about or "data" not in about:
        return None

    profile = about["data"]
    submitted = reddit_get(
        f"https://www.reddit.com/user/{username}/submitted.json?limit={USER_HISTORY_LIMIT}"
    )
    comments = reddit_get(
        f"https://www.reddit.com/user/{username}/comments.json?limit={USER_HISTORY_LIMIT}"
    )

    user_posts = []
    if submitted and "data" in submitted:
        for c in submitted["data"]["children"]:
            d = c["data"]
            user_posts.append({
                "title": d.get("title", "") or "",
                "subreddit": d.get("subreddit", "") or "",
                "selftext": (d.get("selftext", "") or "")[:500],
                "score": d.get("score", 0),
                "created_utc": d.get("created_utc", 0),
            })

    user_comments = []
    if comments and "data" in comments:
        for c in comments["data"]["children"]:
            d = c["data"]
            user_comments.append({
                "body": (d.get("body", "") or "")[:300],
                "subreddit": d.get("subreddit", "") or "",
                "score": d.get("score", 0),
                "created_utc": d.get("created_utc", 0),
            })

    return {
        "username": username,
        "account_created_utc": profile.get("created_utc", 0),
        "total_karma": profile.get("total_karma", 0),
        "posts": user_posts,
        "comments": user_comments,
    }


# =================== LEAD CLASSIFICATION (LLM) ===================

LEAD_SYSTEM_PROMPT = """You are a real estate lead-classifier for posts from Hyderabad real estate subreddits.

Decide whether the post's author is actively in the market to BUY property and would benefit from being contacted by a real estate professional.

A post IS A LEAD if the author is:
- searching for property to buy (flat, villa, plot)
- asking for recommendations on areas, projects, or builders
- comparing options or shortlisting
- asking buying-stage questions where they're clearly the buyer

A post is NOT A LEAD if the author is:
- giving advice or sharing experience after-the-fact
- complaining about a builder/area without buying intent
- sharing news, market analysis, or general discussion
- selling their own property
- looking to rent (not buy)
- asking about commercial leasing only

Output STRICT JSON with EXACTLY these keys (no extras, no markdown):
{
  "is_lead": true | false,
  "reason": "one short sentence explaining the decision",
  "intent_type": "buying" | "investing" | "comparing" | "asking_advice" | "not_a_lead",
  "budget": "string e.g. '1.5cr' or '50-80L' or '' if absent",
  "location": "primary locality/area mentioned, or '' if absent",
  "property_type": "flat" | "villa" | "plot" | "mixed" | "unclear" | "",
  "lead_quality_score": 1-10 integer (10 = ready-to-buy with budget/location/timeline; 1 = vague)
}

If is_lead is false, set lead_quality_score to 0 and intent_type to "not_a_lead".
Return ONLY the JSON object."""


def llm_classify_lead(post):
    user_msg = (
        f"Title: {post['title']}\n"
        f"Flair: {post['flair']}\n"
        f"Body:\n{post['selftext']}"
    )

    resp = anthropic_client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=CLAUDE_MAX_TOKENS,
        system=LEAD_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
        timeout=CLAUDE_TIMEOUT_SEC,
    )

    text = resp.content[0].text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    parsed = json.loads(text)

    parsed["is_lead"] = bool(parsed.get("is_lead", False))
    try:
        parsed["lead_quality_score"] = int(parsed.get("lead_quality_score", 0))
    except (TypeError, ValueError):
        parsed["lead_quality_score"] = 0
    return parsed


# =================== USER CLASSIFICATION (LLM + RULES) ===================

def compute_user_features(history):
    if not history:
        return None

    now = time.time()
    age_days = max(1, int((now - history["account_created_utc"]) / 86400))

    posts = history["posts"]
    comments = history["comments"]

    cutoff_90d = now - (90 * 86400)
    recent_posts = [p for p in posts if p["created_utc"] >= cutoff_90d]
    recent_comments = [c for c in comments if c["created_utc"] >= cutoff_90d]

    sub_counter = Counter(
        [p["subreddit"].lower() for p in recent_posts] +
        [c["subreddit"].lower() for c in recent_comments]
    )
    total = sum(sub_counter.values())

    re_activity = sum(
        v for s, v in sub_counter.items()
        if s in RE_SUBREDDITS or "realestate" in s or "property" in s
    )
    re_pct = round((re_activity / total * 100), 1) if total else 0.0

    promo_hits = 0
    for txt in [p["title"] + " " + p["selftext"] for p in posts] + [c["body"] for c in comments]:
        if not txt:
            continue
        low = txt.lower()
        for pat in PROMO_PATTERNS:
            if re.search(pat, low):
                promo_hits += 1
                break

    top_subs = ", ".join(
        [f"{s}({n})" for s, n in sub_counter.most_common(5)]
    ) or "(none)"

    return {
        "account_age_days": age_days,
        "total_karma": history["total_karma"],
        "posts_90d": len(recent_posts),
        "comments_90d": len(recent_comments),
        "subreddit_diversity": len(sub_counter),
        "re_activity_pct": re_pct,
        "promo_hits": promo_hits,
        "top_subs": top_subs,
    }


USER_TYPE_SYSTEM_PROMPT = """You are classifying Reddit users in Hyderabad real estate subreddits as either an END USER (genuine buyer/researcher) or an AGENT (real estate professional posting commercially).

You will be given:
- mechanical features (account age, karma, % of activity in real estate subs, promotional language hits, subreddit diversity)
- the user's recent post titles and a snippet of recent comments

Signals of an AGENT:
- High % of activity concentrated in real estate subreddits, low diversity
- Promotional language ("DM me", "100% genuine", "premium location", phone numbers)
- Posts pitching properties/projects rather than asking
- Comments that consistently offer help/services
- Repetitive sales-style phrasing across posts

Signals of an END USER:
- Activity spread across diverse subreddits (work, hobbies, city life)
- Posts ask questions about their own situation with personal context
- One-off or short-burst activity in real estate (researching a purchase)
- Conversational tone, not promotional

Output STRICT JSON, no markdown:
{
  "user_type": "end_user" | "agent" | "unclear",
  "confidence": 1-10 integer,
  "reasoning": "one sentence",
  "red_flags": ["short bullet", "..."],
  "supporting_signals": ["short bullet", "..."]
}

Return ONLY JSON."""


def llm_classify_user(features, history):
    if not features or not history:
        return {"user_type": "unclear", "confidence": 0,
                "reasoning": "no user data available",
                "red_flags": [], "supporting_signals": []}

    recent_titles = [p["title"] for p in history["posts"][:5]]
    recent_comments = [c["body"][:200] for c in history["comments"][:5]]

    user_msg = (
        f"MECHANICAL FEATURES:\n"
        f"- account_age_days: {features['account_age_days']}\n"
        f"- total_karma: {features['total_karma']}\n"
        f"- posts_90d: {features['posts_90d']}\n"
        f"- comments_90d: {features['comments_90d']}\n"
        f"- subreddit_diversity_90d: {features['subreddit_diversity']}\n"
        f"- re_activity_pct_90d: {features['re_activity_pct']}\n"
        f"- promo_language_hits: {features['promo_hits']}\n"
        f"- top_subreddits: {features['top_subs']}\n\n"
        f"RECENT POST TITLES:\n" +
        "\n".join(f"- {t}" for t in recent_titles) + "\n\n"
        f"RECENT COMMENT SNIPPETS:\n" +
        "\n".join(f"- {c}" for c in recent_comments)
    )

    resp = anthropic_client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=CLAUDE_MAX_TOKENS,
        system=USER_TYPE_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
        timeout=CLAUDE_TIMEOUT_SEC,
    )

    text = resp.content[0].text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    parsed = json.loads(text)

    parsed.setdefault("user_type", "unclear")
    try:
        parsed["confidence"] = int(parsed.get("confidence", 0))
    except (TypeError, ValueError):
        parsed["confidence"] = 0
    return parsed


# =================== TELEGRAM ===================

TELEGRAM_MAX_CHARS = 3800


def _html_escape(s):
    if not s:
        return ""
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;"))


def _truncate(s, n):
    s = (s or "").strip()
    return s if len(s) <= n else s[: n - 1].rstrip() + "…"


def send_telegram_message(text):
    payload = {
        "chat_id": CHAT_ID, "text": text,
        "parse_mode": "HTML", "disable_web_page_preview": True,
    }
    r = requests.post(TG_SEND_MSG, data=payload, timeout=30)
    r.raise_for_status()


def format_lead_block(idx, post, marker, lead, now_utc):
    posted_dt = datetime.fromtimestamp(post["created_utc"], tz=timezone.utc).astimezone(IST)
    hours_ago = round((now_utc - post["created_utc"]) / 3600, 1)

    title_disp = _truncate(post["title"], 110)
    reason_disp = _truncate(lead.get("reason", ""), 200)
    budget = lead.get("budget", "") or "—"

    title_esc = _html_escape(title_disp)
    reason_esc = _html_escape(reason_disp) or "(no reason)"
    budget_esc = _html_escape(budget)

    marker_str = f" {marker}" if marker else ""

    return (
        f"<b>{idx}.{marker_str} {title_esc}</b>\n"
        f"📅 {posted_dt.strftime('%d-%b %H:%M')} IST · {hours_ago}h ago\n"
        f"💰 {budget_esc} · 🔗 <a href=\"{post['permalink']}\">View on Reddit</a>\n"
        f"📝 {reason_esc}"
    )


def build_messages(leads, now_utc):
    """leads: list of (post, marker, lead_data) tuples."""
    if not leads:
        return []

    end_user_count = sum(1 for _, m, _ in leads if not m)
    unclear_count = sum(1 for _, m, _ in leads if m)

    header_lines = [
        "🎯 <b>New Leads</b> (last check)",
        f"{end_user_count} end_user" + (
            f" + {unclear_count} unclear (high quality) ❓" if unclear_count else ""
        ),
        "",
    ]
    header = "\n".join(header_lines)

    blocks = [
        format_lead_block(i + 1, p, m, l, now_utc)
        for i, (p, m, l) in enumerate(leads)
    ]

    messages = []
    current = header
    for block in blocks:
        if len(current) + len(block) + 2 > TELEGRAM_MAX_CHARS and current.strip() != header.strip():
            messages.append(current.rstrip())
            current = ""
        current += ("\n\n" if current else "") + block

    if current.strip():
        messages.append(current.rstrip())

    return messages


# =================== MAIN ===================

def main():
    now_utc = time.time()
    state = load_state()
    seen_ids = state["seen"]

    posts = fetch_recent_posts()
    if not posts:
        print("No posts fetched.")
        save_state(state)  # still prune
        return

    # Filter to only unseen posts
    new_posts = [p for p in posts if p["id"] and p["id"] not in seen_ids]
    print(f"{len(posts)} posts fetched, {len(new_posts)} are new (not in state).")

    if not new_posts:
        save_state(state)
        return

    # Process each new post: lead classification, user classification, filter
    qualifying_leads = []
    for p in new_posts:
        # 1. Lead detection
        try:
            lead_data = llm_classify_lead(p)
        except Exception as e:
            print(f"[lead-LLM-fail] '{p['title'][:50]}': {e}")
            traceback.print_exc()
            # Mark as seen so we don't retry forever; skip
            seen_ids[p["id"]] = now_utc
            continue

        # Always mark as seen now so we don't reprocess if we crash later
        seen_ids[p["id"]] = now_utc

        if not lead_data.get("is_lead"):
            continue

        # 2. User classification (only for confirmed leads)
        try:
            history = fetch_user_history(p["author"])
        except Exception as e:
            print(f"[user-history-fail] {p['author']}: {e}")
            history = None

        features = compute_user_features(history) if history else None

        try:
            user_class = llm_classify_user(features, history)
        except Exception as e:
            print(f"[user-LLM-fail] {p['author']}: {e}")
            user_class = {"user_type": "unclear", "confidence": 0,
                          "reasoning": "user-classification failed",
                          "red_flags": [], "supporting_signals": []}

        user_type = user_class.get("user_type", "unclear")
        quality = lead_data.get("lead_quality_score", 0)

        # 3. Apply filter: end_user OR (unclear AND quality >= threshold)
        marker = ""
        if user_type == "end_user":
            marker = ""
        elif user_type == "unclear" and quality >= UNCLEAR_QUALITY_THRESHOLD:
            marker = "❓"
        else:
            # agent or low-quality unclear -> skip
            continue

        qualifying_leads.append((p, marker, lead_data))

    # Sort: highest quality first
    qualifying_leads.sort(
        key=lambda x: x[2].get("lead_quality_score", 0),
        reverse=True,
    )

    # Send
    messages = build_messages(qualifying_leads, now_utc)
    for msg in messages:
        send_telegram_message(msg)
        time.sleep(0.4)

    # Persist state
    save_state(state)

    print(
        f"Done. New posts processed: {len(new_posts)}. "
        f"Qualifying leads sent: {len(qualifying_leads)}. "
        f"Messages sent: {len(messages)}."
    )


if __name__ == "__main__":
    main()