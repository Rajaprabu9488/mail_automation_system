from langchain_core.prompts import ChatPromptTemplate
from langchain_ollama import ChatOllama
import json


def mail_reply_generator(user_role, message_subject, message_content):
    llm = ChatOllama(
        model="llama3.2:3b",
        temperature=0.2,
        keep_alive=-1
    )

    prompt = ChatPromptTemplate.from_template("""
      You are an AI email reply assistant.

      Your task is to generate exactly three possible email replies to the received email.

      USER ROLE:
      {user_role}

      ORIGINAL EMAIL SUBJECT:
      {subject}

      ORIGINAL EMAIL CONTENT:
      {email_content}

      Generate exactly these three reply variants:

      1. POSITIVE
         A polite and professional reply that accepts, agrees with, or positively responds to the email.

      2. NEGATIVE
         A polite and professional reply that declines, disagrees with, or cannot accept the request.

      3. MAYBE LATER
         A polite and professional reply that shows interest but requests more time, postponement, or an alternative arrangement.

      IMPORTANT RULES:

      * Consider the USER ROLE when writing the replies.
      * Understand the context of the original email before generating the responses.
      * The user's role is the only role provided. Do not ask for or invent the sender's role.
      * Generate a separate subject for each reply.
      * Subjects must be short, clear, professional, and relevant to the original email.
      * Do not simply copy the original subject unless it is appropriate.
      * Every email body MUST begin with exactly:

      Dear Sir/Madam,

      * Use a formal, polite, and professional tone.
      * Keep the email concise and natural.
      * Use simple and clear English.
      * Do not use overly complicated words.
      * Do not use casual expressions or emojis.
      * Do not invent people's names.
      * Never use placeholders such as [Name], [Your Name], [Interviewer's Name], [Company Name], etc.
      * Do not ask the user for missing information.
      * Do not invent dates, times, meetings, commitments, reasons, or other facts that are not present in the original email.
      * Use only information available in the original email and the user's role.
      * End every email body with exactly:

      Best regards

      * Do not include explanations or reasoning.
      * Do not include Markdown.
      * Do not include code fences.
      * Return ONLY valid JSON.
      * Use double quotes for all JSON keys and string values.
      * Do not add any text before or after the JSON.
      * Escape quotation marks inside email content when necessary.

      OUTPUT FORMAT:

      {{
      "positive": {{
      "subject": "Short professional subject",
      "body": "Dear Sir/Madam,\n\nConcise professional reply.\n\nBest regards"
      }},
      "negative": {{
      "subject": "Short professional subject",
      "body": "Dear Sir/Madam,\n\nConcise professional reply.\n\nBest regards"
      }},
      "maybe_later": {{
      "subject": "Short professional subject",
      "body": "Dear Sir/Madam,\n\nConcise professional reply.\n\nBest regards"
      }}
      }}
   """)

    chain = prompt | llm

    result = chain.invoke({
        "user_role": user_role,
        "subject": message_subject,
        "email_content": message_content
    })

    return json.loads(result.content)