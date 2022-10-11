import docx

if __name__ == '__main__':
    doc = docx.Document('C:/Users/РДС/Desktop/Test.docx')  # Открытие документа
    all_str = doc.paragraphs  # Разбор документа на строки

    new_file = open("file.txt", "w+")  # Создание текстового файла

    count = 0  # Переменная счетчика строки

    for para in all_str:  # Проход по строкам файла
        if para.style == doc.styles['Normal']:  # Определение строки вопроса
            if count == 0:  # Определение первой и последующих строк вопроса
                new_file.write('::' + para.text + '{\n')
            else:
                new_file.write('}\n\n::' + para.text + '{\n')
            count += 1
        elif para.style == doc.styles['List Paragraph']:  # Определение строки ответа
            if para.text.find('*') > 0:  # Определение правильного ответа по "*"
                new_file.write('\t=' + para.text + '\n')
            else:
                new_file.write('\t~' + para.text + '\n')
    new_file.write('}')  # Окончание обработки файла
