import boto3

key_id = 'your_access_key'
secret_key_id = 'your_secret_key'

session = boto3.session.Session()
s3 = session.client(
    service_name='s3',
    endpoint_url='https://storage.yandexcloud.net',
    aws_access_key_id=key_id,
    aws_secret_access_key=secret_key_id
)

bucket_name = 'bucket-name'
file_key = 'uploads/uvao_ng_ticket_yyyymmddHHMM.xlsx'

def check_file_in_bucket(bucket_name, file_key):
    try:
        s3.head_object(Bucket=bucket_name, Key=file_key)
        print(f"Файл '{file_key}' найден в бакете '{bucket_name}'.")
    except s3.exceptions.ClientError as e:
        if e.response['Error']['Code'] == "404":
            print(f"Файл '{file_key}' не найден в бакете '{bucket_name}'.")
        else:
            print(f"Ошибка при проверке файла: {e}")

check_file_in_bucket(bucket_name, file_key)
