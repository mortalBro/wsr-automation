import mysql.connector
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from countticket import countTicketReport
from fourthpage import fourthPage
from insertionInstance import insertionOfInstanceCount
from thirdpage import thirdPageWork
from thirteenpage import thirteenpageWork
from utils import *



presentation = Presentation('/home/tspl/Documents/wsrAutomation/WSR (Weekly Status Report).pptx')

print("first page work start")
current_monday=current_monday()
previous_monday=previous_monday()
current_sunday=current_sunday()
previous_sunday=previous_sunday()
firstPageChange(presentation,previous_monday,current_monday)
firstPageChange(presentation,previous_sunday,current_sunday)
print("first page work done")

print("Instance count csv Updation only for Mirror")
res=countTicketReport()
insertionOfInstanceCount(res)
print("Instance count csv Updation only for Mirror is Completed",res)
print("third page executive Summary updation Start")
thirdPageWork()
print("third page executive Summary updation completed")
print("Fourth page executive Summary updation start")
fourthPage()
print("Fourth page executive Summary updation completed")
print("thirteen page executive Summary updation start")
thirteenpageWork()
print("thirteen page executive Summary updation completed")

