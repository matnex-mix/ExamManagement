from django.db.models import Q
from django.shortcuts import render,redirect
from django.utils import timezone
import os
from .forms import LoginForm, UploadResultForm,HodViewForm,ExamOfficerViewForm,chatForm
from .models import UserTable,ResultTable,messageTable,communicatingPartyId
from win32com.client import Dispatch
import pythoncom
from . import worksheet
# Create your views here.
def Login(request):
    if request.method == 'POST':
        form = LoginForm(request.POST)
        if form.is_valid():
            UserId = request.POST['UserId']
            Passcode = request.POST['Passcode']
            try:
                user = UserTable.objects.get(UserId = UserId, Passcode = Passcode)
            except:
                context = {'form':form, 'error': 'Unknow Login Credentials'}
                return render(request, 'login.html', context)
            if user.UserType == 'Lecturer':
                request.session['Lecturer'] = UserId
                return redirect('LecturerHomepage')
            elif user.UserType == 'HOD':
                request.session['HOD'] = UserId
                return redirect('HodHomepage')
            else:
                request.session['ExamOfficer'] = UserId
                return redirect('ExamOfficerHomepage')

    else:
        form = LoginForm()
    return render(request, 'login.html',{'form':form})

def LecturerHomepage(request):
    user = request.session.get('Lecturer')
    if user == None:
        return redirect('LoginPage')
    return render(request,'Lecturer/homepage.html')

def UploadResult(request):
    user = request.session.get('Lecturer')
    if user == None:
        return redirect('LoginPage')
    if request.method == 'POST':
        form = UploadResultForm(request.POST,request.FILES)
        form.changeCourseCode('Default', 'Default')

        if form.is_valid():
            Class = request.POST['Class']
            Session = request.POST['Session']
            Semester = request.POST['Semester']
            Units = request.POST['Units']
            CourseCode = ''
            file = ''
            try:
                CourseCode = request.POST['CourseCode']
            except:
                pass
            if CourseCode == '':
                form.changeCourseCode(Class, Semester)
                return render(request, 'Lecturer/Upload.html', {'form': form, 'error': 'Please Select  course Code'})
            elif CourseCode != '':
                try:
                     file = request.FILES['ExamFile']
                except:
                     return render(request, 'Lecturer/Upload.html', {'form': form,'error':'Exam Sheet is required'})
            DateUploaded = timezone.now()
            check = ''
            try:
                check = ResultTable.objects.get(Class = Class,Session = Session, Semester = Semester, CourseCode = CourseCode)
            except:
                pass
            if check != '':
                return render(request, 'Lecturer/Upload.html', {'form': form, 'error': 'Result already Submitted'})
            insert = ResultTable(LecturerId = user,
                                 Class = Class,
                                 Session = Session,
                                 Semester = Semester,
                                 CourseCode = CourseCode,
                                 Units = Units,
                                 ExamSheet = file,
                                 DateUploaded = DateUploaded)
            insert.save()
            form = UploadResultForm()
            return render(request, 'Lecturer/Upload.html', {'form': form, 'error': 'Exam Sheet successfully Submitted'})
        else:
            form.changeCourseCode('Invalid', 'Invalid')

    else:
        form = UploadResultForm()
    return render(request,'Lecturer/Upload.html',{'form':form})

def LecturerViewPending(request):
    user = request.session.get('Lecturer')
    if user == None:
        return redirect('LoginPage')
    PendingResult = ResultTable.objects.filter(LecturerId = user, AcceptedByHod = 'No')
    PendingResult = PendingResult.order_by('-id')
    return render(request,'Lecturer/PendingResult.html',{'Pending': PendingResult})

def LecturerEditResult(request,ResultId):
    user = request.session.get('Lecturer')
    if user == None:
        return redirect('LoginPage')
    pythoncom.CoInitialize()
    try:
        getFile = ResultTable.objects.get(id = ResultId,LecturerId = user)
    except:
        return redirect('LecturerHomepage')
    fileOpen = Dispatch("Excel.Application")
    fileOpen.Visible = True
    ExcelFile = 'C:/Users/LORDFM/Desktop/python/Aishat/examManagement'+ getFile.ExamSheet.url
    if getFile.AcceptedByHod == 'No':
        ExcelFile = fileOpen.Workbooks.Open(str(ExcelFile))
    else:
        ExcelFile = fileOpen.Workbooks.Open(str(ExcelFile), None, True)
    return render(request,'Lecturer/EditResult.html')


def LecturerAllAcceptedResults(request):
    user = request.session.get('Lecturer')
    if user == None:
        return redirect('LoginPage')
    AllResults = ResultTable.objects.filter(LecturerId = user, AcceptedByHod = 'Yes')
    AllResults = AllResults.order_by('-id')
    return render(request,'Lecturer/AllResults.html',{'AllResults': AllResults})


def HodHomepage(request):
    user = request.session.get('HOD')
    if user == None:
        return redirect('LoginPage')
    return render(request,'Hod/homepage.html')

def HodViewResultCategory(request):
    user = request.session.get('HOD')
    if user == None:
        return redirect('LoginPage')

    form = HodViewForm()
    return render(request,'Hod/ResultCategory.html',{'form': form})


def HodViewResult(request):
    user = request.session.get('HOD')
    if user == None:
        return redirect('LoginPage')
    if request.method == 'POST':
        form = HodViewForm(request.POST)
        if form.is_valid():
            Class = request.POST['Class']
            Semester = request.POST['Semester']
            Category = request.POST['Category']
            AllResults =''
            if Category == 'Pending':
                AllResults = ResultTable.objects.filter(Class = Class,Semester = Semester, AcceptedByHod = 'No')
            else:
                AllResults = ResultTable.objects.filter(Class=Class, Semester=Semester, AcceptedByHod='Yes')
            AllResults = AllResults.order_by('-id')
    else:
        return redirect('HodSelectResultCategoryPage')
    return render(request,'Hod/AllResults.html',{'AllResults':AllResults})


def HodViewResultDetails(request,ResultId):
    user = request.session.get('HOD')
    if user == None:
        return redirect('LoginPage')
    pythoncom.CoInitialize()
    try:
        getFile = ResultTable.objects.get(id = ResultId)
    except:
        return redirect('HodSelectResultCategoryPage')
    fileOpen = Dispatch("Excel.Application")
    fileOpen.Visible = True
    ExcelFile = 'C:/Users/LORDFM/Desktop/python/Aishat/examManagement'+ getFile.ExamSheet.url
    ExcelFile = fileOpen.Workbooks.Open(str(ExcelFile), None, True)
    return render(request,'Hod/ViewResult.html')

def HodForwardToExamOficer(request,ResultId):
    user = request.session.get('HOD')
    if user == None:
        return redirect('LoginPage')
    if request.method =='POST':
        get = ''
        try:
            get = ResultTable.objects.get(id = ResultId, AcceptedByHod = 'No')
        except:
            return redirect('HodSelectResultCategoryPage')
        get.AcceptedByHod = 'Yes'
        get.save()
        return render(request,'Hod/Forward.html',{'success': 'Forwarded'})
    return render(request,'Hod/Forward.html')

def ExamOfficerHomepage(request):
    user = request.session.get('ExamOfficer')
    if user == None:
        return redirect('LoginPage')
    return render(request,'ExamOfficer/homepage.html')


def ExamOfficerViewResultCategory(request):
    user = request.session.get('ExamOfficer')
    if user == None:
        return redirect('LoginPage')

    form = ExamOfficerViewForm()
    return render(request,'ExamOfficer/ResultCategory.html',{'form': form})


def ExamOfficerViewResult(request):
    user = request.session.get('ExamOfficer')
    if user == None:
        return redirect('LoginPage')
    AllResults = ''
    if request.method == 'POST':
        form = ExamOfficerViewForm(request.POST)
        if form.is_valid():

            Class = request.POST['Class']
            Semester = request.POST['Semester']
            Category = request.POST['Category']
            if Category == 'Pending':
                AllResults = ResultTable.objects.filter(Class=Class, Semester=Semester, AcceptedByHod='Yes', Compiled = 'No')
            else:
                AllResults = ResultTable.objects.filter(Class=Class, Semester=Semester, AcceptedByHod='Yes',Compiled='Yes')
            AllResults = AllResults.order_by('-id')
        else:
            return render(request,'ExamOfficer/ResultCategory.html',{'form': form})
    else:
        return redirect('ExamOfficerSelectResultCategoryPage')
    return render(request,'ExamOfficer/AllResults.html',{'AllResults':AllResults})


def ExamOfficerViewResultDetails(request,ResultId):
    user = request.session.get('ExamOfficer')
    if user == None:
        return redirect('LoginPage')
    pythoncom.CoInitialize()
    try:
        getFile = ResultTable.objects.get(id = ResultId, Compiled = 'No')
    except:
        return redirect('ExamOfficerSelectResultCategoryPage')
    fileOpen = Dispatch("Excel.Application")
    fileOpen.Visible = True
    ExcelFile = 'C:/Users/LORDFM/Desktop/python/Aishat/examManagement'+ getFile.ExamSheet.url
    ExcelFile = fileOpen.Workbooks.Open(str(ExcelFile), None, True)
    return render(request,'ExamOfficer/ViewResult.html')

def ExamOfficerViewCompiledResult(request,ResultId):
    user = request.session.get('ExamOfficer')
    if user == None:
        return redirect('LoginPage')
    pythoncom.CoInitialize()
    try:
        getFile = ResultTable.objects.get(id = ResultId, Compiled = 'Yes')
    except:
        return redirect('ExamOfficerSelectResultCategoryPage')
    fileOpen = Dispatch("Excel.Application")
    fileOpen.Visible = True
    ExcelFile = getFile.CompiledFile
    ExcelFile = fileOpen.Workbooks.Open(str(ExcelFile), None, True)
    return render(request,'ExamOfficer/ViewResult.html')


def ExamOfficerCompile(request,ResultId):
    user = request.session.get('ExamOfficer')
    if user == None:
        return redirect('LoginPage')
    if request.method =='POST':
        get = ''
        try:
            get = ResultTable.objects.get(id = ResultId, Compiled = 'No')
        except:
            return redirect('ExamOfficerSelectResultCategoryPage')
        Session = get.Session
        ValueSessionToFunction = Session
        Session = str.replace(Session, ' ', '')
        Session = str.replace(Session, '/', '')

        Class = get.Class
        Semester = get.Semester
        Units = get.Units
        CourseCode = get.CourseCode
        CourseCode = str.replace(CourseCode,' ','')
        if Semester == 'First Semester':
            Semester = '01'
        else:
            Semester = '02'
        Directory = str.replace(os.getcwd(),"\\","/")
        InitialSheet =Directory+get.ExamSheet.url
        CompiledFileName = Directory+'/media/compiled/'+Class + Session + Semester + '.xlsx'
        worksheet.CompileResult(IndividualSheet= InitialSheet,CompiledSheet=CompiledFileName,CourseCode=CourseCode,CourseUnit=Units,
                                Session= ValueSessionToFunction)
        get.Compiled = 'Yes'
        get.CompiledFile = CompiledFileName
        get.save()

        return render(request,'ExamOfficer/Compiled.html',{'success': 'Compiled'})
    return render(request,'ExamOfficer/Compiled.html')


def ExamOfficerViewAllUsers(request):
    a = allUsers(request,'ExamOfficer')
    return a

def ExamOfficerViewUserDetails(request,userId):
    a = profile(request,'ExamOfficer', userId)
    return a

def ExamOfficerChat(request,chatId):
    a = chat(request,chatId,'ExamOfficer')
    return a

def ExamOfficerAllMessages(request):
    a = allChats(request,'ExamOfficer')
    return a

#lecturer chat starts here
def LecturerViewAllUsers(request):
    a = allUsers(request,'Lecturer')
    return a

def LecturerViewUserDetails(request,userId):
    a = profile(request,'Lecturer', userId)
    return a

def LecturerChat(request,chatId):
    a = chat(request,chatId,'Lecturer')
    return a

def LecturerAllMessages(request):
    a = allChats(request,'Lecturer')
    return a

#Hod chat starts here
def HodViewAllUsers(request):
    a = allUsers(request,'HOD')
    return a

def HodViewUserDetails(request,userId):
    a = profile(request,'HOD', userId)
    return a

def HodChat(request,chatId):
    a = chat(request,chatId,'HOD')
    return a

def HodAllMessages(request):
    a = allChats(request,'HOD')
    return a


"""These sections contains code chatting"""
def allUsers(request,userType):
    user = request.session.get(userType)
    if user == None:
        return redirect('LoginPage')

    usersList = UserTable.objects.all()
    usersList = usersList.exclude(UserId = user)
    if userType == 'ExamOfficer':
        return render(request,'ExamOfficer/allUsers.html',{'users': usersList,'user':user})
    elif userType == 'Lecturer':
        print (user)
        return render(request,'Lecturer/allUsers.html',{'users': usersList,'user':user})
    else:
        return render(request,'Hod/allUsers.html',{'users': usersList,'user':user})
def profile(request, userType, userId):
    user = request.session.get(userType)
    if user == None:
        return redirect('LoginPage')
    getUuserDetails = UserTable.objects.get(id = userId)
    comParty_1 = user + ' '+ getUuserDetails.UserId
    comParty_2 = getUuserDetails.UserId +' '+ user
    comId = ''
    try:
        check = communicatingPartyId.objects.get(Q(parties = comParty_1) | Q(parties = comParty_2))
        comId = check.id
    except:
        comId = communicatingPartyId(parties = comParty_1)
        comId.save()
        comId = comId.id

    context = {'userDetails': getUuserDetails,'comId':comId}
    if userType == 'ExamOfficer':
        return render(request, 'ExamOfficer/userDetails.html', context)
    elif userType == 'Lecturer':
        return render(request, 'Lecturer/userDetails.html', context)
    else:
        return render(request, 'Hod/userDetails.html', context)


# chat View
def chat(request,chatId,userType):
    user = request.session.get(userType)
    print(userType)
    if user == None:
        return redirect('LoginPage')
    chatIdentifier = ''
    lastChat = ''
    try:
        chatIdentifier = communicatingPartyId.objects.get(id = chatId, parties__contains = user)
    except:
        print ('error')
        return redirect('LoginPage')
    secondParty = chatIdentifier.parties
    secondParty = str.split(secondParty, ' ')
    chatHistory = messageTable.objects.filter(chatId = chatId)
    if len(chatHistory) != 0 :
        lastChat = chatHistory.last()
        if lastChat.receiver == user:
            lastChat.status = 'Seen'
            lastChat.save()
    if user != secondParty[0]:
        secondParty = secondParty[0]
    else:
        secondParty = secondParty[1]
    if request.method == 'POST':
        form = chatForm(request.POST)
        if form.is_valid():
            message = request.POST['message']
            dateSent = timezone.now()
            if len(chatHistory) != 0 :
                lastChat.lastMessage = 'No'
                lastChat.save()
            insert = messageTable(chatId = chatId,
                                    message = message,
                                    sender = user,
                                    receiver = secondParty,
                                    dateSent = dateSent,
                                    status = 'Unseen',
                                    lastMessage = 'Yes')
            insert.save()
            if userType == 'ExamOfficer':
                return redirect('ExamOfficerChatPage',chatId)
            elif userType == 'Lecturer':
                return redirect('LecturerChatPage', chatId)
            else:
                return redirect('HodChatPage', chatId)
    else:
        form = chatForm()
    context ={'chatHistory':chatHistory, 'form' : form, 'owner': user}
    if userType == 'ExamOfficer':
        return render(request,'ExamOfficer/chats.html', context)
    elif userType == 'Lecturer':
        return render(request,'Lecturer/chats.html', context)
    else:
        return render(request,'Hod/chats.html', context)


# all messages View
def allChats(request,userType):
    user = request.session.get(userType)
    if user == None:
        return redirect('LoginPage')
    getMessages = messageTable.objects.filter((Q(receiver = user) | Q( sender = user)) & Q (lastMessage = 'Yes'))
    getMessages = getMessages.order_by('-id')
    objectReturned = []
    for f in getMessages:
        otherParty = ''
        if (f.sender == user):
            otherParty = f.receiver
        else:
            otherParty = f.sender
        secondParty = otherParty
        if secondParty == 'admin':
            if len(f.message) > 100:
                f.message = f.message[0: 100]
            objectReturned.append({'chatId':f.chatId,'message': f.message,'otherParty': secondParty,'dateSent':f.dateSent
                                   ,'status':f.status,'sender': f.sender})
        else:

            otherParty = UserTable.objects.get(UserId = otherParty)
            surname = otherParty.Surname
            firstname = otherParty.Firstname
            names = surname +' '+ firstname

            if len(f.message) > 100:
                f.message = f.message[0: 100]
            objectReturned.append({'chatId':f.chatId,'message': f.message,'otherParty': secondParty,'dateSent':f.dateSent
                                   ,'status':f.status,'image': otherParty,'names':names,'sender': f.sender})
    if userType == 'ExamOfficer':
        return render(request, 'ExamOfficer/messages.html',{'allMessages': objectReturned,'owner':user})
    elif userType == 'Lecturer':
        return render(request, 'Lecturer/messages.html',{'allMessages': objectReturned,'owner':user})
    else:
        return render(request, 'Hod/messages.html', {'allMessages': objectReturned, 'owner': user})