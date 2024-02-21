from typing import Optional
from pydantic import BaseModel, EmailStr
import config, random


class CertificateIn1(BaseModel):
    name: str
    certification_type : str = random.choice(config.certification_types)
    achivement_type : str = random.choice(config.achivement_types)
    mentor_name : str = random.choice(config.mentor_names)
    mentor_type : str = random.choice(config.mentor_types)
    insititue_name : str = random.choice(config.insititue_names) 
    course_name : str = random.choice(config.course_names)
    mail : Optional[EmailStr] = None


class CertificateIn2(CertificateIn1):
    director_name : str = random.choice(config.director_names)
    director_type : str = random.choice(config.director_types)


class CertificateOut1(CertificateIn1):
    url : str


class CertificateOut2(CertificateIn2):
    url : str
