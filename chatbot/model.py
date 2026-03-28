import os
import requests
from dotenv import load_dotenv

load_dotenv()

# ════════════════════════════════════════════════
#  CONFIG
# ════════════════════════════════════════════════

api_key=os.getenv("GROQ_API_KEY")   # https://console.groq.com
GROQ_MODEL   = "llama-3.1-8b-instant"

SYSTEM_PROMPT = """You are StudentAI, an intelligent academic assistant designed to help students learn effectively.

Your role:
- Answer questions in a clear, academic, and student-friendly way
- Give a direct answer first, then explanation, then a real-world example if helpful
- For math/science: show step-by-step reasoning
- For history/literature: give context and significance  
- Keep answers concise but complete — 3-6 sentences normally, more if the topic needs it
- Never say "As an AI..." — just answer directly
- If you don't know something, say so honestly

Tone: Encouraging, clear, like a knowledgeable older student helping a peer."""

# ════════════════════════════════════════════════
#  SMALL TALK  (instant, no API call wasted)
# ════════════════════════════════════════════════

import random

SMALL_TALK = {
    "hi":        ["Hello! What are you studying today?", "Hi there! What can I help you with?"],
    "hello":     ["Hello! Ask me anything academic.", "Hi! What topic do you need help with?"],
    "hey":       ["Hey! What subject are we tackling?", "Hey! What's on your mind?"],
    "bye":       ["Goodbye! Keep studying hard!", "See you! Good luck with your studies."],
    "goodbye":   ["Goodbye! Come back whenever you need help."],
    "thanks":    ["You're welcome! Ask me anything else.", "Happy to help!"],
    "thank you": ["You're welcome! Good luck!", "Anytime!"],
    "ok":        ["Great! Ask me anything whenever you're ready."],
    "okay":      ["Sounds good! What would you like to learn?"],
}

def check_small_talk(text):
    cleaned = text.lower().strip().rstrip("!.,?")
    return random.choice(SMALL_TALK[cleaned]) if cleaned in SMALL_TALK else None

# ════════════════════════════════════════════════
#  GROQ API CALL
# ════════════════════════════════════════════════

def ask_groq(user_input):
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type":  "application/json"
    }
    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_input}
        ],
        "temperature": 0.5,
        "max_tokens":  400,
        "top_p":       0.9
    }
    try:
        response = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=15
        )
        if response.status_code == 200:
            reply = response.json()["choices"][0]["message"]["content"].strip()
            print(f"[Groq] OK — {len(reply)} chars")
            return reply
        else:
            print(f"[Groq] Error {response.status_code}: {response.text[:300]}")
            return None
    except requests.exceptions.Timeout:
        print("[Groq] Request timed out")
        return None
    except Exception as e:
        print(f"[Groq] Exception: {e}")
        return None

# ════════════════════════════════════════════════
#  MAIN FUNCTION  (called by Flask)
# ════════════════════════════════════════════════

def chatbot_response(user_input, mode="student"):
    user_input = user_input.strip()

    if not user_input:
        return "Please type or say something — I'm here to help!"

    # Small talk gets instant reply, no API call
    small_talk_reply = check_small_talk(user_input)
    if small_talk_reply:
        return small_talk_reply

    # Everything else goes directly to Groq
    reply = ask_groq(user_input)

    if reply:
        return reply

    # Only reached if API completely fails
    return "I'm having trouble connecting right now. Please check your internet or try again in a moment."