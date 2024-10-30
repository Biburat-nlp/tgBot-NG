from datetime import datetime
import boto3
import os
import time
from pyarrow.ipc import new_file
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, DirCreatedEvent, FileCreatedEvent

session = boto3.session.Session()
s3 = session.client(
    service_name='s3',
    endpoint_url='https://storage.yandexcloud.net'
)

bucket_name = 'name'
local_directory = 'local'

class S3UploadHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return None

        file_path = event.src_path
        if file_path.endswith('.xlsx'):
            timestamp = datetime.now().strftime('%Y%m%d%H%M')
            new_file_name = f"uvao_ng_ticket_{timestamp}.xlsx"

            try:
                s3.upload_file(file_path, bucket_name, f'uploads/{new_file_name}')
                print(f"Файл {new_file_name} загружен в стораге.")
            except Exception as e:
                print(f"Ошибка при загрузке файла {new_file_name}, Альберт насрал в код: {e}")

event_handler = S3UploadHandler()
observer = Observer()
observer.schedule(event_handler, path=local_directory, recursive=False)
observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()
observer.join()