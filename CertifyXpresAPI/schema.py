from pydantic import BaseModel
import config, random

class Certificate(BaseModel):
    name: str
    certification_type : str = random.choice(config.certification_types)
    achivement_type : str = random.choice(config.achivement_types)
    mentor_name : str = random.choice(config.mentor_names)
    mentor_type : str = random.choice(config.mentor_types)
    insititue_name : str = random.choice(config.insititue_names) 
    course_name : str = random.choice(config.course_names)

class CertificateOut(Certificate):
    url : str