from fastapi import FastAPI, Query, status, UploadFile, File 
import config, s3config, schema, os, openpyxl
from util import replace_text_in_docx, convert_to_pdf
from datetime import datetime
from fastapi.responses import JSONResponse


app = FastAPI()

@app.get("/")
async def root():
    return {"message": "This is the root of the CertifyXpresAPI, please use the /docs endpoint to see the API documentation."}


@app.post("/c/",status_code=status.HTTP_201_CREATED)
async def create_certification(certificate: schema.Certificate,
                            template: str = Query(..., title="Template", description="Select a template", enum=config.template_options)):
        

       
        certificate_template = config.CERTIFICATE_TEMPLATE_PATH + f'{template}'
        print(f"Selected template: {certificate_template}")
        
        doc_output_path = fr'{config.DOC_OUTPUT_PATH}/{certificate.name}.docx'
        pdf_output_path = fr'{config.PDF_OUTPUT_PATH}/{certificate.name}.pdf'

        await replace_text_in_docx(
        template_path = certificate_template, 
        certificate = certificate, 
        output_path=doc_output_path
        )
        await convert_to_pdf(doc_output_path, pdf_output_path)

        # Ask user if they want to upload the certificate to S3

        s3_client = s3config.S3Config()
        url = s3_client.upload_certificate(
            pdf_output_path,
            client_name=certificate.name, 
            certificate_name=certificate.name)
        print(f"Certificate uploaded to {url}")

        print(f"PDF saved to {pdf_output_path}")
        print("Certificate generated successfully!")
    
        return {
              "message": "Certificate generated successfully!",
              "local path": pdf_output_path,
              "Online url": url}


# route to create multiple certificate through excel
@app.post("/c/excel", status_code=status.HTTP_201_CREATED)
async def create_certification_from_excel(
    files: UploadFile = File(..., title="Excel file", description="Upload an excel file"),
    template: str = Query(..., title="Template", description="Select a template", enum=config.template_options)
):
    try:
        # Create a unique name for the uploaded file
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        file_name = f"{timestamp}_{files.filename}"
        excel_file_path = os.path.join(config.EXCEL_ROOT_PATH, file_name)

        # Save the uploaded Excel file
        with open(excel_file_path, "wb") as buffer:
            buffer.write(await files.read())

        wb = openpyxl.load_workbook(excel_file_path)
        sheet = wb.active

        certificate_template = os.path.join(config.CERTIFICATE_TEMPLATE_PATH, template)
        print(f"Selected template: {certificate_template}")

        name = "shubhamXYZ"

        # Update doc_path to a unique name with timestamp
        # Update doc_output_path to a unique name with timestamp
        doc_output_path = os.path.join(config.DOC_OUTPUT_PATH, f"{name}-{timestamp}.docx")

        # iterate through the rows in the excel file
        for row in sheet.iter_rows(values_only=True):
            ename = row[0]

            await replace_text_in_docx(
                template_path=certificate_template,
                certificate=schema.Certificate(name=ename),
                output_path=doc_output_path
            )

            pdf_output_path = os.path.join(config.PDF_OUTPUT_PATH, name, f"{ename}.pdf")

            if not os.path.exists(os.path.join(config.PDF_OUTPUT_PATH, name)):
                os.makedirs(os.path.join(config.PDF_OUTPUT_PATH, name))

            try:
                await convert_to_pdf(doc_output_path, pdf_output_path)

                s3_client = s3config.S3Config()
                url = s3_client.upload_certificate(
                    pdf_output_path,
                    client_name=name,
                    certificate_name=ename)
                print(f"Certificate saved & uploaded to {url}")

                print(f"PDF saved to {pdf_output_path}")
            except Exception as e:
                print(f"An error occurred during PDF conversion: {e}")

            print(f"Certificate generated for {ename}")
            print(f"PDF saved to {pdf_output_path}")

        print("Certificates generated successfully!")
        return JSONResponse(content={"message": "Certificates generated successfully!"}, status_code=status.HTTP_201_CREATED)

    except Exception as e:
        return JSONResponse(content={"message": str(e)}, status_code=status.HTTP_500_INTERNAL_SERVER_ERROR)
