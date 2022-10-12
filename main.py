import docx
import easygui as g


def translate_docx():
    file_name = g.fileopenbox('Выберите файл')  # Выбор документа

    doc = docx.Document(file_name)

    all_str = doc.paragraphs  # Разбор документа на строки

    new_file = open(file_name.replace('docx', 'txt'), "w+")  # Создание текстового файла

    count = 0  # Переменная счетчика строки

    for para in all_str:  # Проход по строкам файла
        if para.style == doc.styles['Normal']:  # Определение строки вопроса
            if count == 0:  # Определение первой и последующих строк вопроса
                new_file.write('::' + ' '.join(para.text.split()[:3]) + '... ::' + para.text + '{\n')
            else:
                new_file.write('}\n\n::' + ' '.join(para.text.split()[:3]) + '... ::' + para.text + '{\n')
            count += 1
        elif para.style == doc.styles['List Paragraph']:  # Определение строки ответа
            if para.text.find('*') > 0:  # Определение правильного ответа по "*"
                new_file.write('\t=' + para.text.replace('*', '') + '\n')
            else:
                new_file.write('\t~' + para.text + '\n')

    new_file.write('}')  # Окончание обработки файла

    print('Выполнение завершено!')


if __name__ == '__main__':
    translate_docx()
