from nltk.tokenize import sent_tokenize
import requests

def translate(text, target_lang, auth_key):
    url = 'https://api-free.deepl.com/v2/translate'
    headers = {
        'Authorization': f'DeepL-Auth-Key {auth_key}'
    }
    data = {
        'text': text,
        'target_lang': target_lang
    }

    response = requests.post(url, headers=headers, data=data)

    if response.status_code == 200:
        result = response.json()
        translations = result.get('translations', [])
        if translations:
            translated_text = translations[0].get('text')
            detected_source_language = translations[0].get('detected_source_language')
            return translated_text, detected_source_language
        else:
            return None, None
    else:
        print(f"Translation failed with status code {response.status_code}")
        return None, None


auth_key = '592206a4-2b37-704a-6031-646fea11502a:fx'
target_language = 'KO'

# 번역할 문장 입력 후 함수에 전달
with open('data.txt', 'r', encoding='utf-8') as file:
    text = file.read()

sentences = sent_tokenize(text)

translated_sentences = []
for sentence in sentences:
    translated_sentences.append(translate(sentence))

for sentence in sentences:
    print(sentence + '\n')

for translated_sentence in translated_sentences:
    print(translated_sentence + '\n')

for index in range(len(sentences)):
    print(sentences[index] + '\n' + translated_sentences[index] + '\n')

