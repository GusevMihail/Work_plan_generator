import smtplib  # Импортируем библиотеку по работе с SMTP
import os  # Функции для работы с операционной системой, не зависящие от используемой операционной системы

# Добавляем необходимые подклассы - MIME-типы
import mimetypes  # Импорт класса для обработки неизвестных MIME-типов, базирующихся на расширении файла
from email import encoders  # Импортируем энкодер
from email.mime.base import MIMEBase  # Общий тип
from email.mime.text import MIMEText  # Текст/HTML
from email.mime.image import MIMEImage  # Изображения
from email.mime.audio import MIMEAudio  # Аудио
from email.mime.multipart import MIMEMultipart  # Многокомпонентный объект


def send_email(addr_to, msg_subj, msg_text, files):
    addr_from = 'kzs.asu.504@gmail.com'  # Отправитель
    password = 'kzsasu123456'  # Пароль

    msg = MIMEMultipart()  # Создаем сообщение
    msg['From'] = addr_from  # Адресат
    msg['To'] = addr_to  # Получатель
    msg['Subject'] = msg_subj  # Тема сообщения

    body = msg_text  # Текст сообщения
    msg.attach(MIMEText(body, 'plain'))  # Добавляем в сообщение текст

    process_attachement(msg, files)

    # ======== Этот блок настраивается для каждого почтового провайдера отдельно ===============================================
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)  # Создаем объект SMTP
    server.ehlo()
    # server.starttls()                                      # Начинаем шифрованный обмен по TLS
    # server.set_debuglevel(True)                            # Включаем режим отладки, если не нужен - можно закомментировать
    server.login(addr_from, password)  # Получаем доступ
    server.send_message(msg)  # Отправляем сообщение
    server.quit()  # Выходим
    # ==========================================================================================================================


def process_attachement(msg, files):  # Функция по обработке списка, добавляемых к сообщению файлов
    for f in files:
        if os.path.isfile(f):  # Если файл существует
            attach_file(msg, f)  # Добавляем файл к сообщению
        elif os.path.exists(f):  # Если путь не файл и существует, значит - папка
            dir = os.listdir(f)  # Получаем список файлов в папке
            for file in dir:  # Перебираем все файлы и...
                attach_file(msg, f + "/" + file)  # ...добавляем каждый файл к сообщению
        else:
            raise Exception(f'File or directory not found: {f}')


def attach_file(msg, filepath):  # Функция по добавлению конкретного файла к сообщению
    filename = os.path.basename(filepath)  # Получаем только имя файла
    ctype, encoding = mimetypes.guess_type(filepath)  # Определяем тип файла на основе его расширения
    if ctype is None or encoding is not None:  # Если тип файла не определяется
        ctype = 'application/octet-stream'  # Будем использовать общий тип
    maintype, subtype = ctype.split('/', 1)  # Получаем тип и подтип
    if maintype == 'text':  # Если текстовый файл
        with open(filepath) as fp:  # Открываем файл для чтения
            file = MIMEText(fp.read(), _subtype=subtype)  # Используем тип MIMEText
            fp.close()  # После использования файл обязательно нужно закрыть
    elif maintype == 'image':  # Если изображение
        with open(filepath, 'rb') as fp:
            file = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
    elif maintype == 'audio':  # Если аудио
        with open(filepath, 'rb') as fp:
            file = MIMEAudio(fp.read(), _subtype=subtype)
            fp.close()
    else:  # Неизвестный тип файла
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)  # Используем общий MIME-тип
            file.set_payload(fp.read())  # Добавляем содержимое общего типа (полезную нагрузку)
            fp.close()
            encoders.encode_base64(file)  # Содержимое должно кодироваться как Base64
    file.add_header('Content-Disposition', 'attachment', filename=filename)  # Добавляем заголовки
    msg.attach(file)  # Присоединяем файл к сообщению


def send_journals(batch: dict, attachment_folder: str, mail_subj: str, add_month_to_subj: bool = True,
                  subj_suffix: str = '',
                  mail_text: str = '', print_log=True, test_mod=False):
    from application import get_xlsx_files
    import config_email
    # folder = r'./output data/journals/'
    all_journals = tuple(get_xlsx_files(attachment_folder))
    mail_subj = f'{mail_subj} {get_month_str(all_journals[0]) if add_month_to_subj else ""} {subj_suffix}'

    if test_mod:
        print('Test mode: emails will not sending')
    if print_log:
        print(f'email subject: {mail_subj}')

    for addr, journals_aliases in batch.items():
        files_to_send = []
        for file_name in all_journals:
            for j in journals_aliases:
                if j in file_name:
                    files_to_send.append(attachment_folder + file_name)
                    break

        if print_log:
            print(f'{addr.ljust(20)} {files_to_send}')
        if not test_mod:
            send_email(addr, mail_subj, mail_text, files_to_send)


def get_month_str(attachment_name: str) -> str:
    from datetime import datetime
    import locale
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
    date_string = attachment_name.split('/')[-1][0:7]  # + ' 01'
    month = datetime.strptime(date_string, '%Y %m').strftime('%B')
    return month
