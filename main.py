import csv

from django.shortcuts import render

from .models import RollFile
from .convertor import GUser, GUser_school


def GsuiteConvertor(request):
    template_name = "g-suite_convertor.html"
    context = {}

    if request.method == "POST":
        file = request.FILES["roll_file"]
        school = request.POST.get("school")
        s_info = valid_G(school, file)
        if s_info:
            grade = request.POST.get("grade")
            classN = request.POST.get("classN")
            file_name = f"{school}{grade}-{classN}_user.csv"
            roll_file = RollFile(title=file_name, roll_file=file)
            roll_file.roll_file.name = file_name
            if request.POST.get("whole_school"):
                guser = GUser_school(file, s_info)
            else:
                guser = GUser(file, s_info, grade, classN)
            roll_file.save()
            context["result"] = guser.file_url
        else:
            context["errors"] = "학교명 혹은 파일이 올바르지 않습니다."
    return render(request, template_name, context)


def valid_G(school, file):
    s_info = ""
    with open("staticfiles/G-suite/cbe_school_info.csv", "r", newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        for row in reader:
            if row[1] == school:
                s_info = [row[0], school, row[2].split("@")[0]]
                break

    if s_info:
        # Should change .csv to .xlsx when deploy
        if file.size < 100000 and ".xlsx" in file.name:
            return s_info
        print(f"{file.name}: {file.size}byte")

    return False
