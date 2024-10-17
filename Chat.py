import json
from difflib import get_close_matches
import win32com.client
import os
import speech_recognition as sr
speaker = win32com.client.Dispatch("SAPI.SpVoice")
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing .....")
            user_input = r.recognize_google(audio, language="en-in")
            return user_input
        except Exception as e:
            return "Some error had occcured please come back in sometime"
def speak(text):
    os.system(f"say{text}")
def load_knowledge_base(file_path: str) -> dict:
    with open(file_path, 'r') as file:
        data: dict = json.load(file)
    return data
def save_knowledge_base(file_path: str, data:dict):
    with open(file_path, 'w') as file:
        json.dump(data, file, indent=2)
def find_best_match(user_question: str, questions: list[str]) -> str | None:
    matches: list=get_close_matches(user_question, questions, n=1, cutoff=0.6)
    return matches[0] if matches else None
def get_answer_for_question(question: str, knowledge_base:dict) -> str | None:
    for q in knowledge_base["questions"]:
        if q["question"] == question:
            return q["answer"]
def chat_bot():
    knowledge_base: dict = load_knowledge_base('knowledge_base.json')
    while True:
        print('listinning... ')
        user_input = takecommand()
        if user_input.lower() =='break':
            break
        best_match: str | None = find_best_match(user_input, [q["question"] for q in knowledge_base["questions"]])
        if best_match:
            answer: str = get_answer_for_question(best_match, knowledge_base)
            speaker.speak(f' {answer}')
        else:
            speaker.speak(' I dont know the answer. Can you type the answer')
            new_answer: str = input('Type the answer or "skip" to skip: ')
            if new_answer.lower() != 'skip':
                knowledge_base["questions"].append({"question": user_input, "answer": new_answer})
                save_knowledge_base('knowledge_base.json',knowledge_base)
                speaker.speak(' Thank you! I learned new response!')
