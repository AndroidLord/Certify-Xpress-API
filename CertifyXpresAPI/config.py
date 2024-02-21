import os

CERTIFICATE_TEMPLATE_PATH='C:/Users/PC/Documents/Experiments/Certification Template/'

DOC_OUTPUT_PATH = fr'C:\Users\PC\Documents\Experiments\Certificates\\'
PDF_OUTPUT_PATH = fr'C:\Users\PC\Documents\Experiments\Pdf'
EXCEL_ROOT_PATH = fr'C:\Users\PC\Documents\Experiments\excel/'


template_options = [f for f in os.listdir(CERTIFICATE_TEMPLATE_PATH) if f.endswith('.docx')]


certification_types = ["Software Development", "Data Science", "Web Development", "Network Security", "Machine Learning"]
achivement_types = ["First Place", "Top Performer", "Outstanding Achievement", "Excellence Award", "Innovator Award"]
insititue_names = ["Tech Institute", "Data Science Academy", "Web Development School", "Cybersecurity Center", "AI Research Institute"]
course_names = ["Python Programming", "Data Analysis with Python", "Full-Stack Web Development", "Network Security Fundamentals", "Machine Learning Essentials"]
mentor_names = ["Dr. Smith", "Ms. Johnson", "Mr. Williams", "Prof. Brown", "Dr. Garcia"]
mentor_types = ["Technical Mentor", "Career Mentor", "Project Mentor", "Academic Mentor", "Industry Mentor"]
director_names = ["Dr. Smith", "Ms. Johnson", "Mr. Williams", "Prof. Brown", "Dr. Garcia"]
director_types = ["Technical Director", "Career Director", "Project Director", "Academic Director", "Industry Director"]
