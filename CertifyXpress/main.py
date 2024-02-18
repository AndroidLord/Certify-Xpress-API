import config, s3config
from util import replace_text_in_docx, convert_to_pdf
import os, openpyxl
from datetime import datetime
import asyncio



async def main():
      # Prompt user for details
    
    # ask if this is a single or multiple certificate generation
    print("Welcome to CertifyXpress!")
    
    name = input("Enter the name of the recipient: ")
     # Output paths for the generated certificate
    doc_output_path = fr'{config.DOC_OUTPUT_PATH}/{name}.docx'
    pdf_output_path = fr'{config.PDF_OUTPUT_PATH}/{name}.pdf'
        
    # Path to the certificate template
    certificate_template = config.CERTIFICATE_TEMPLATE_PATH + 'Sample-Certificate-Template.docx'


    # ask if this is a single or multiple certificate generation
    print("If you want to generate a single certificate, enter '1'")
    print("If you want to generate multiple certificates, enter '2'")

    # get user input
    user_input = input("Enter your choice: ")

    # if user wants to generate a single certificate
    if user_input == '1':

        print("Hey, Do you also it want to upload the certificate to S3? y/n")
        upload_to_s3 = input("Enter your choice: ")

        await replace_text_in_docx(
        template_path = certificate_template, 
        to_replace='name',
        replace_with = name, 
        output_path=doc_output_path
        )
        await convert_to_pdf(doc_output_path, pdf_output_path)

        # Ask user if they want to upload the certificate to S3

        if upload_to_s3 == 'y':
            s3_client = s3config.S3Config()
            url = s3_client.upload_certificate(
                pdf_output_path,
                client_name=name, 
                certificate_name=name)
            print(f"Certificate uploaded to {url}")
        else:
            print(f"PDF saved to {pdf_output_path}")
            print("Certificate generated successfully!")

    # if user wants to generate multiple certificates
    elif user_input == '2':

        print("Hey, Do you also it want all the certificates to be uploaded to S3? y/n")
        upload_to_s3 = input("Enter your choice: ")


        # get the list of names
        excel_file_path = config.EXCEL_ROOT_PATH + 'Book1.xlsx'
        wb = openpyxl.load_workbook(excel_file_path)
        sheet = wb.active

        # Get current time in milliseconds
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")

        # Update doc_path to a unique name with timestamp
        # Update doc_output_path to a unique name with timestamp
        doc_output_path = fr'{config.DOC_OUTPUT_PATH}/{name}-{timestamp}.docx'

        replacing_text = 'name'
        # iterate through the rows in the excel file
        for row in sheet.iter_rows(values_only=True):
            ename = row[0]
            
            await replace_text_in_docx(
                template_path = certificate_template, 
                to_replace='name',
                replace_with = ename, 
                output_path=doc_output_path
            )

            replacing_text = ename

            pdf_output_path = config.PDF_OUTPUT_PATH + fr'\{name}\{ename}.pdf'
            
            if not os.path.exists(config.PDF_OUTPUT_PATH + fr'\{name}'):
                os.makedirs(config.PDF_OUTPUT_PATH + fr'\{name}')

            try:
                await convert_to_pdf(doc_output_path, pdf_output_path)
                if upload_to_s3 == 'y':
                    s3_client = s3config.S3Config()
                    url = s3_client.upload_certificate(
                        pdf_output_path,
                        client_name=name, 
                        certificate_name=ename)
                    print(f"Certificate saved & uploaded to {url}")
                else:
                    print(f"PDF saved to {pdf_output_path}")
            except Exception as e:
                print(f"An error occurred during PDF conversion: {e}")
            
            print(f"Certificate generated for {ename}")
            print(f"PDF saved to {pdf_output_path}")

            
        print("Certificates generated successfully!")

if __name__ == "__main__":
    asyncio.run(main())
