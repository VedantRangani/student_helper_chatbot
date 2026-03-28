from chatbot.model import chatbot_response

while True:
    msg = input("You: ")
    if msg.lower() == "exit":
        break
    print("Bot:", chatbot_response(msg))

