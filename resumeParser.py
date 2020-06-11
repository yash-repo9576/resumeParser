import docx2txt
import spacy
from spacy.matcher import Matcher
import re
import pandas as pd
from nltk.corpus import stopwords
import PyPDF2

nlp = spacy.load('en_core_web_sm')
STOPWORDS = set(stopwords.words('english'))

# Education Degrees
EDUCATION = [
            'BE','B.E.', 'B.E', 'BS', 'B.S',
            'ME', 'M.E', 'M.E.', 'MS', 'M.S',
            'BTECH', 'B.TECH', 'M.TECH', 'MTECH',
            'SSC', 'HSC', 'CBSE', 'ICSE', 'X', 'XII'
        ]


# extracting text from pdf files
def extract_text_from_pdf(pdf_path):
    # creating a pdf file object
     pdfFileObj = open(pdf_path, 'rb')
    # creating a pdf reader object
     pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
     # extracting text from page
     pageObj = pdfReader.getPage(0)
     temp = pageObj.extractText()
     text = [line.replace('\t', ' ') for line in temp.split('\n') if line ]
     #print(pageObj.extractText())
     return ' '.join(text)


# extracting text from docx files
def extract_text_from_doc(doc_path):
    temp = docx2txt.process(doc_path)
    text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
    return ' '.join(text)


# returns name as a string
def extract_name(resume_text):
    matcher = Matcher(nlp.vocab)
    pattern = [{'POS': 'PROPN'}, {'POS': 'PROPN'}]
    matcher.add('NAME', None, pattern)

    nlp_text = nlp(resume_text)
    matches = matcher(nlp_text)

    for match_id, start, end in matches:
        span = nlp_text[start:end]
        return span.text


# returns number as a string
def extract_number(resume_text):
    phone_number = re.findall(re.compile(
        r'(?:(?:\+?([1-9]|[0-9][0-9]|[0-9][0-9][0-9])\s*(?:[.-]\s*)?)?(?:\(\s*([2-9]1[02-9]|[2-9][02-8]1|[2-9][02-8][02-9])\s*\)|([0-9][1-9]|[0-9]1[02-9]|[2-9][02-8]1|[2-9][02-8][02-9]))\s*(?:[.-]\s*)?)?([2-9]1[02-9]|[2-9][02-9]1|[2-9][02-9]{2})\s*(?:[.-]\s*)?([0-9]{4})(?:\s*(?:#|x\.?|ext\.?|extension)\s*(\d+))?'),
                              resume_text)
    if phone_number:
        number = ''.join(phone_number[0])
        if len(phone_number) > 10:
            return '+' + number
        else:
            return number


# returns email as a string
def extract_email(resume_text):
    pattern = re.compile(r'[a-zA-Z0-9|\.]+@[a-zA-Z|\.]+\.[a-z|\.]+')
    email = re.findall(pattern, resume_text)
    if email:
        try:
            return email[0]
        except IndexError:
            return None


# returns skills as a list
def extract_skills(resume_text):
    nlp_text = nlp(resume_text)
    noun_chunks = nlp_text.noun_chunks

    tokens = [token.text for token in nlp_text if not token.is_stop]

    data = pd.read_csv("skills.csv")

    skills = list(data.columns.values)

    skillset = []

    for token in tokens:
        if token.lower() in skills:
            skillset.append(token)

    for token in noun_chunks:
        token = token.text.lower().strip()
        if token in skills:
            skillset.append(token)

    return [i.capitalize() for i in set([i.lower() for i in skillset])]

# returns education as a tuple
def extract_education(resume_text):
    nlp_text = nlp(resume_text)

    # Sentence Tokenizer
    nlp_text = [sent.string.strip() for sent in nlp_text.sents]

    edu = {}
    # Extract education degree
    for index, text in enumerate(nlp_text):
        for tex in text.split():
            # Replace all special symbols
            tex = re.sub(r'[?|$|.|!|,]', r'', tex)
            if tex.upper() in EDUCATION and tex not in STOPWORDS:
                edu[tex] = text + nlp_text[index + 1]

    # Extract year
    education = []
    for key in edu.keys():
        year = re.search(re.compile(r'(((20|19)(\d{2})))'), edu[key])
        if year:
            education.append((key, ''.join(year[0])))
        else:
            education.append(key)
    return education


Resume_text = extract_text_from_doc('/home/phoenix/Resume.docx')
# Resume_text = extract_text_from_pdf('/home/phoenix/Resume.pdf')

print('Name : ' + extract_name(Resume_text))
print('Contact Number : ' + extract_number(Resume_text))
print('Email : ' + extract_email(Resume_text))
print('Skills : ')
for skill in extract_skills(Resume_text):
    print(skill)
print('Education : ')
for edu in extract_education(Resume_text):
    print(edu)
