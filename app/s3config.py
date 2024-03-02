import os
import boto3
from dotenv import load_dotenv

class SingletonMeta(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            instance = super().__call__(*args, **kwargs)
            cls._instances[cls] = instance
        return cls._instances[cls]

class S3Config(metaclass=SingletonMeta):
    def __init__(self):
        load_dotenv()
        self.s3_client = boto3.client(
            service_name=os.getenv('SERVICE_NAME'),
            region_name=os.getenv('REGION_NAME'),
            aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
            aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY')
        )
        self.bucket_name = os.getenv('BUCKET_NAME')

    def list_objects(self, bucket_name, folder_prefix):
        response = self.s3_client.list_objects_v2(Bucket=bucket_name, Prefix=folder_prefix)
        return response


    # certificate_path = fr'C:\Users\PC\Documents\Experiments\Pdf\shubham\jake.pdf'
    

    def upload_certificate(self,certificate_path,client_name,certificate_name):
        print('Uploading certificate to S3...')
        # Set the content type explicitly to 'application/pdf'
        extra_args = {'ContentType': 'application/pdf'}
        self.s3_client.upload_file(
            certificate_path,
            self.bucket_name,
            fr'certificates\{client_name}\{certificate_name}.pdf',
            ExtraArgs=extra_args
        )

        # return the URL of the certificate
        return self.s3_client.generate_presigned_url('get_object', Params={'Bucket': self.bucket_name, 'Key': fr'certificates\{client_name}\{certificate_name}.pdf'})


    # # run upload_certificate function
    # url = upload_certificate(certificate_path)
    # print(f"Certificate uploaded to {url}")