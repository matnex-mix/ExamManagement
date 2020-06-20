from django.db import models

# Create your models here.
class UserTable(models.Model):
    UserId = models.CharField(max_length=15)
    UserType = models.CharField(max_length=15)
    Surname = models.CharField(max_length=15)
    Firstname = models.CharField(max_length=15)
    Passcode = models.CharField(max_length=20)

class ResultTable(models.Model):
    LecturerId = models.CharField(max_length=20)
    Class = models.CharField(max_length=12)
    Session = models.CharField(max_length=11)
    Semester = models.CharField(max_length=15)
    CourseCode = models.CharField(max_length=7)
    Units = models.IntegerField()
    ExamSheet = models.FileField(upload_to='result/')
    DateUploaded = models.DateTimeField()
    AcceptedByHod = models.CharField(max_length=3,default='No')
    Compiled = models.CharField(max_length=3,default='No')
    CompiledFile = models.CharField(max_length=150, null= True)

class messageTable (models.Model):
    chatId = models.CharField(max_length=100)
    message = models.CharField(max_length=3000)
    sender = models.CharField(max_length=15)
    receiver = models.CharField(max_length=15)
    dateSent = models.DateTimeField()
    status = models.CharField(max_length=10)
    lastMessage = models.CharField(max_length=3)


class communicatingPartyId(models.Model):
    parties = models.CharField(max_length=50)
