import docx
import easygui as g
from docx.shared import RGBColor


def change_file():
    name = g.fileopenbox('Выберите файл')  # Выбор документа
    if name.find('docx'):
        translate_docx(name)  # Вызов функции обработки файла
    else:
        print('Вы выбрали некорректный файл!')


def translate_docx(file_name):
    doc = docx.Document(file_name)  # Открытие файла

    flag_for_question = False

    new_file = open(file_name.replace('docx', 'txt'), "w+")  # Создание текстового файла

    count = 0  # Переменная счетчика строки

    for para in doc.paragraphs:  # Проход по строкам файла

        if para.text.find('Вопрос #') != 0 and para.text.find('Ссылка на НТД') != 0 and \
                not ('ФНП Правила' in para.text) and not ('Указаний по' in para.text) and not ('п. ' in para.text):
            if para.runs[0].bold:
                flag_for_question = False
                if count == 0:
                    new_file.write('::' + ' '.join(para.text.split()[:3]) + '... ::' + para.text + '{\n')
                else:
                    new_file.write('}\n\n::' + ' '.join(para.text.split()[:3]) + '... ::' + para.text + '{\n')
                count += 1
            elif para.runs[0].font.color.rgb == RGBColor(33, 150, 83):
                if flag_for_question:
                    new_file.write('\t~%50%' + para.text + '\n')
                else:
                    new_file.write('\t=' + para.text + '\n')
            else:
                if para.text == '*Может быть несколько верных вариантов':
                    flag_for_question = True
                else:
                    new_file.write('\t~' + para.text + '\n')

    new_file.write('}')  # Окончание обработки файла


if __name__ == '__main__':
    change_file()
