import json
import requests
import pandas as pd

courseID = [2027062,
1814256,
1811513,
1811514,
1823120,
1984342,
1792643,
1970312,
2069731,
2170149,
1824226,
1811600,
1824383,
2277019,
2218771,
1822269,
1793091,
1976682,
1841387,
2082789,
1793094,
1988827,
1793098,
2017485,
2136192,
2247139,
1811822,
2019607,
2019614,
2164788,
1792675,
2126556,
1998326,
2145207,
1991305,
2056248,
2102704,
2110215,
2066067,
2113073,
2156089,
2126555,
1823733,
2114338,
2025735,
2098729,
2096095,
2096102,
2098732,
2098766,
2096098,
2096101,
2096109,
2098756,
2096108,
2096091,
2096096,
2098720,
2046829,
2137931,
2099253,
2048114,
2128391,
2127801,
2170497,
2110016,
2006081,
2085843,
2085849,
2226861,
2097220,
1837892,
2117270,
2019685,
1989849,
2117689,
1931627,
2075456,
2109102,
2165488,
2069214,
2052504,
2165655,
2081189,
2165652,
2166455,
2148888,
2148896,
2048084,
2052500]

course_name = [
'(Archived) Abstract Mathematics',
'(Archived) AMC Algebra Preparation',
'(Archived) AP Art History',
'(Archived) AP Calculus AB',
'(Archived) AP Calculus BC',
'(Archived) AP Chinese',
'(Archived) AP Computer Science A',
'(Archived) AP English Literature',
'(Archived) AP European History',
'(Archived) AP Macroeconomics',
'(Archived) AP Microeconomics',
'(Archived) AP Physics C: Electricity and Magnetism',
'(Archived) Ordinary Differential Equations',
'AP Biology',
'AP Environmental Science',
'AP French Language and Culture',
'AP Government',
'AP Human Geography',
'AP Music Theory',
'AP Physics 1',
'AP Physics 2',
'AP Psychology',
'AP Statistics',
'AP U.S. History',
'AP World History: Modern',
'British Literature II, (1750-Present)',
'BT5 Instructor Course',
'Calculus III',
'Chinese I',
'Chinese II',
'Cognitive Psychology',
'College Admissions 2',
'Computational Astrophysics',
'Computer Vision',
'COVID-19 Data Analysis',
'Deep Learning',
'Development Economics',
'EKG Technician',
'Engineering Design I',
'English Composition',
'Exoplanets and Extraterrestrial Life',
'Exploring The Intersection: Neurodiversity and Psychology',
'Fundamentals of Astronomy and Astrophysics',
'General Microbiology',
'Genetics',
'Human Anatomy',
'IGCSE Additional Maths',
'IGCSE Biology',
'IGCSE Business Studies',
'IGCSE Chemistry',
'IGCSE Computer Science',
'IGCSE Economics',
'IGCSE English Language',
'IGCSE English Literature',
'IGCSE Extended Mathematics',
'IGCSE French',
'IGCSE Global Perspectives',
'IGCSE Hindi',
'IGCSE Physics',
'Intro to Biochemistry',
'Introduction to Business Management',
'Introduction to Capitalism',
'Introduction to COVID-19 Research',
'Introduction to Philosophy',
'Introduction to Self-Care',
'Introduction to the Supreme Court',
'Introduction to Video Game History',
'Latin I',
'Latin II',
'Medical Terminology',
'Modern Physics',
'Music History',
'Personal Finance and Financial Literacy',
'Postclassical World History: 476-800 AD',
'Principles of Engineering',
'Probability Theory',
'Python 2: GUI Development',
'SAT Prep',
'Theoretical Physics',
'AMC Geometry Preparation',
'AP Spanish Language and Culture',
'Applications of Neuroscience',
'Arabic I',
'Biotechnology',
'Graph Theory',
'Introduction to Photoshop and Lightroom',
'Italian I',
'Italian II',
'Neuroscience I',
'Neuroscience II',
]
    
headers = {
    'Authorization': 'Bearer 7~56QzXqgAb3YqNI9bohaxX9fTxWQDKDPwC3vM1xZqe5hAMlg89oYxaZiibT1VSuu7',
}

for course in courseID:
    response = requests.get(f'https://canvas.instructure.com/api/v1/courses/{course}/analytics/student_summaries', headers=headers)
    data = response.json() # convert to a JSON format so you can actually work with the data
    
    n = courseID.index(course)
    filename = f"{course_name[n]}.json"

    with open(filename, 'w') as out_file:
        json.dump(data, out_file)



for course in course_name:
    try:
        filename = f'C:\\Users\\duyth\\Documents\\Beyond The Five Audits\\Audit Data\\{course}.json'
        df = pd.read_json(filename)
        df.to_csv(f'C:\\Users\\duyth\\Documents\\Beyond The Five Audits\\Audit Data\\{course}.csv', index=None)
    
        csv_filename = f'C:\\Users\\duyth\\Documents\\Beyond The Five Audits\\Audit Data\\{course}.csv'
        x = pd.read_csv(csv_filename)
        x.to_excel(f'C:\\Users\\duyth\\Documents\\Beyond The Five Audits\\Audit Data\\{course}.xlsx')
    except:
        print(f'{course} json file is corrupted')


newCourseName = []
for course in course_name:
    if len(course) >= 26:    
        new = course[0:26]
        newCourseName.append(new)
    else:
        newCourseName.append(course)

xls = pd.ExcelFile('C:\\Users\\duyth\\OneDrive\\Documents\\Beyond The Five Audits\\Data.xlsx')


for course in newCourseName:
    try: 
        df = pd.read_excel(xls, course)
        df.drop(columns=["page_views", 'max_page_views', 'page_views_level', 'participations', 'max_participations', 'participations_level'], inplace = True)
        df.to_excel(f'C:\\Users\\duyth\\OneDrive\\Documents\\Beyond The Five Audits\\{course}.xlsx')
    except:
        print(f'{course} is corrupted')