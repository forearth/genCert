from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
import os
from csv import reader, DictReader

# 수료증 생성 함수
def genCert(title, no, name, course, organization, period, times, issueDate):
    documentTitle = title
    no=no
    studentName=name
    fileName =f'{no}({studentName}).pdf'
    courseName=course
    organizationName=organization
    period=period
    times=times
    issueDate=issueDate

    # 수료내용
    textLines = [
    f'위 사람은 {organizationName}에서 실시한',
    f'"{courseName}" 과정을',
    '이수하였으므로 이 증서를 수여합니다.'
    ]
    logoImg = 'logo.jpg'

    # 라이브러리 불러오기 및 설정
    saveName=os.path.join(".\\certs\\", fileName)
    pdf = canvas.Canvas(saveName)
    pdf.setTitle(documentTitle)

    pdfmetrics.registerFont(
        TTFont('nanumMJBold', 'NanumMyeongjoExtraBold.ttf')
    )
    pdfmetrics.registerFont(
        TTFont('nanumMJ', 'NanumMyeongjoBold.ttf')
    )

    pdf.setFont('nanumMJBold', 62)
    pdf.drawCentredString(300, 630, documentTitle)

    pdf.setFont("nanumMJ", 20)

    # 제목과 내용 정리
    no=f'제{no}호'
    studentName=f'성      명 : {studentName}'
    courseName=f'과 정 명  : {courseName}'
    period=f'기      간 : {period}'
    times=f'교육시간 : {times}'

    # 일련번호
    pdf.drawString(70,720, no)
    # 성명
    pdf.drawString(70,540, studentName)
    # 과정명
    pdf.drawString(70,510, courseName)
    # 기간
    pdf.drawString(70,480, period)
    # 교육시간
    pdf.drawString(70,450, times)

    # 발행일자
    pdf.setFont("nanumMJ", 16)
    pdf.drawCentredString(300, 240, issueDate)

    # 발행기관
    pdf.setFont("nanumMJBold", 36)
    pdf.drawCentredString(300, 120, organizationName)
    
    # 수료내용
    text = pdf.beginText(70, 380)
    text.setFont("nanumMJBold", 28)
    text.setFillColor(colors.black)
    for line in textLines:
        text.setLeading(40)
        text.textLine(line)

    pdf.drawText(text)

    # 로고 이미지 삽입
    pdf.drawInlineImage(logoImg, 240, 160, width=120, height=30)

    # 최종저장
    pdf.save()


# 수료학생정보 및 수료과정 정보가 담긴 students.csv를 한줄씩 불러오면서 수료증 생성
with open('students.csv', 'r') as read_obj:
    csv_dict_reader = DictReader(read_obj)
    for row in csv_dict_reader:
        title=row['title']
        no=row['no']
        name=row['name']
        course=row['course']
        org=row['organization']
        period=row['period']
        times=row['times']
        issueDate=row['issue date']
        genCert(title, no, name, course, org, period, times, issueDate)