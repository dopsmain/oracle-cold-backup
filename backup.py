# -*- coding: utf8 -*-
import zipfile
import os
import time
import wmi
import logging
import sys
import shutil

########## MAIN CONFIGURATION
service_name = 'Service Name' #Имя сервиса Oracle
date = time.strftime('%d.%m.%Y') #определение времени создания
date_time = time.strftime('%d.%m.%Y_%H.%M.%S') #определение времени создания
########## END MAIN CONFIGURATION

########## PATH CONFIGURATION
PROJECT_PATH = os.path.abspath(os.path.dirname(__file__))

path_archive_dir_DB = os.path.join(PROJECT_PATH, "DB/")
path_archive_dir_LOG = os.path.join(PROJECT_PATH, "LOG/")

file_backup_db = os.path.join(path_archive_dir_DB, "%s_%s_DB.zip") % (date_time, service_name,)
file_backup_log = os.path.join(path_archive_dir_LOG, "%s_%s_LOG.zip") % (date_time, service_name,)

path_src_dir_DB = r'C:\oracle\oradata\DB_NAME'
path_src_dir_LOG = r'C:\oracle\oradata\LOG'

share_server = r'\\server\oracle_backup'
########## END PATH CONFIGURATION

########## LOG CONFIGURATION
path_log = os.path.join(PROJECT_PATH, 'logs/%s_%s.log') % (service_name, date_time,)
logging.basicConfig(filename=path_log,format='%(asctime)s %(levelname)s %(message)s',level=logging.DEBUG)

# Возможные сообщения
# logging.debug('This is a debug message')
# logging.info('This is an info message')
# logging.warning('This is a warning message')
# logging.error('This is an error message')
# logging.critical('This is a critical error message')
########## END LOG CONFIGURATION

########## SERVICE
def service_manager(action, service):
    """
     Управление Windows службами запуск, остановка
     Передается 2 параметра:
        action  - какую задачу выполнять. Может быть "stop", "start"
        service - имя службы, не путаете с "Выводимым именем" службы, они могут быть разными
    """
    c = wmi.WMI()
    for service in c.Win32_Service(Name=service):
        if action == 'stop':
            if service.State == 'Running':
                result, = service.StopService()
                if result == 0:
                    srvc_msg_00 = u"Служба %s остановлена" % service.Name
                    logging.info(srvc_msg_00)
                else:
                    srvc_msg_01 = u"!!!При остановке службы возникли проблемы!!!"
                    logging.error(srvc_msg_01)
            else:
                logging.warning(u'Служба не запущена')
            break
        elif action == 'start':
            if service.State == 'Running':
                logging.warning(u'Служба уже была запущена')
            else:
                result, = service.StartService()
                if result == 0:
                    srvc_msg_02 = u"Служба %s запущена" % service.Name
                    logging.info(srvc_msg_02)
                else:
                    srvc_msg_03 = u"!!!При старте службы возникли проблемы!!!"
                    logging.error(srvc_msg_03)
            break
    else:
        srvc_msg_04 = u"Служба не найдена"
        logging.error(srvc_msg_04)
########## END SERVICE

########## MAKE ARHIVE
def zip_folder(folder_path, output_path):
    """
    Архивация указанно папки со всеми подпапками(включая пустые) в ZIP архив.
    Передается 2 параметра:
        folder_path - исходная папка, которую необходимо архивировать
        output_path - имя архива с полным путями например:
            D:\oracle_backup\LOG\01.03.2012_14.33.14_LOG.zip
    """
    parent_folder = os.path.dirname(folder_path)
    # путь к папке
    contents = os.walk(folder_path)
    try:
        zip_file = zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED, allowZip64 = True)
        for root, folders, files in contents:
            # Добавить все подпапки, включая пустые.
            for folder_name in folders:
                absolute_path = os.path.join(root, folder_name)
                relative_path = absolute_path.replace(parent_folder + '\\','')
                # Логируем процесс
                add_arv_msg_00 = u"Файл %s добавлен в архив" % absolute_path
                logging.info(add_arv_msg_00)
                # Запись архива
                zip_file.write(absolute_path, relative_path)
            for file_name in files:
                absolute_path = os.path.join(root, file_name)
                relative_path = absolute_path.replace(parent_folder + '\\','')
                # Логируем процесс
                add_arv_msg_01 = u"Файл %s добавлен в архив." % absolute_path
                logging.info(add_arv_msg_01)
                # Запись архива
                zip_file.write(absolute_path, relative_path)
        add_arv_msg_02 = u"Архив %s успешно создан." % output_path.replace(PROJECT_PATH,'')
        logging.info(add_arv_msg_02)
    except (IOError, OSError, zipfile.BadZipfile, zipfile.LargeZipFile), message:
        logging.error(message)
        sys.exit(1)
    finally:
        zip_file.close()
########## END MAKE ARHIVE

########## DEL OLD ORACLE LOG
def del_log_file(folder):
    for the_file in os.listdir(folder):
        file_path = os.path.join(folder, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
                del_file_msg_01 = u"Файл %s удален." % the_file
                logging.info(del_file_msg_01)
        except Exception, message:
            logging.error(message)
########## END DEL OLD ORACLE LOG

########## RUN
if __name__ == '__main__':
    # Остановка службы
    service_manager("stop", service_name)
    # Архивация DB
    zip_folder(path_src_dir_DB, file_backup_db)
    # Архивация LOG
    zip_folder(path_src_dir_LOG, file_backup_log)
    # Чистка логов
    del_log_file(path_src_dir_LOG)
    # Перенос DB архива на другой сервер, с логированием
    try:
        shutil.move(file_backup_db, share_server)
        logging.info(u'Файл %s перемещен в %s' % (file_backup_db.replace(path_archive_dir_DB,''), share_server,))
    except (OSError, IOError, shutil.Error), message:
        logging.error(message)
    # Перенос LOG архива на другой сервер, с логированием
    try:
        shutil.move(file_backup_log, share_server)
        logging.info(u'Файл %s перемещен в %s' % (file_backup_log.replace(path_archive_dir_LOG,''), share_server,))
    except (OSError, IOError, shutil.Error), message:
        logging.error(message)
    # Запуск службы
    service_manager("start", service_name)
########## END RUN