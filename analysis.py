import pandas as pd

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

newCourseName = []
for course in course_name:
    if len(course) >= 26:    
        new = course[0:26]
        newCourseName.append(new)
    else:
        newCourseName.append(course)

Stats = {}

for course in newCourseName:
    try:
        df = pd.read_excel(f'C:\\Users\\duyth\\OneDrive\\Documents\\Beyond The Five Audits\\Audit Data\\{course}.xlsx')

        on_time = df['tardiness_breakdown__on_time'].sum()
        total = df['tardiness_breakdown__total'].sum()
        Proportion = on_time/total

        course_stats = {course: [on_time, total, Proportion]}
        Stats.update(course_stats)
    except:
        print(f'{course} data is not available')

summary_stats = pd.DataFrame(Stats)
summary_stats.to_excel('C:\\Users\\duyth\\OneDrive\\Documents\\Beyond The Five Audits\\export.xlsx')