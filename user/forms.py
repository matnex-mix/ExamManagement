from django import forms
from . import Choices
class LoginForm(forms.Form):
        UserId = forms.CharField(max_length=15,
                                 widget=forms.TextInput(attrs={'class': 'textField', 'placeholder': 'Your Id here'}),
                                 label='User Id')
        Passcode = forms.CharField(
            widget=forms.PasswordInput(attrs={'class': 'textField', 'placeholder': 'Your Password here'}),
            max_length=20,
            label='Password')

class UploadResultForm(forms.Form):
    Class = forms.ChoiceField(widget=forms.Select({'class':'selectField','id':'Class'}),choices=Choices.ClassChoices())
    Units = forms.ChoiceField(widget=forms.Select({'class': 'selectField', 'id': 'Units'}), choices=Choices.UnitChoice())
    Session = forms.ChoiceField(widget=forms.Select({'class':'selectField','id':'Session'}),choices=Choices.SessionChoice())

    Semester = forms.ChoiceField(widget=forms.Select({'class':'selectField','id':'Semester'}),choices=Choices.SemesterChoice())

    CourseCode = forms.ChoiceField(widget=forms.Select({'class':'selectField','id':'CourseCode'})
                                   ,required=False)
    ExamFile = forms.FileField(label='Exam Sheet', required=False)

    def changeCourseCode(self,Class,Session):

        if Session == 'First Semester' and Class == 'HND1':
            self.fields['CourseCode'].choices = Choices.HND1FirstSemester()
        elif Session == 'Second Semester' and Class == 'HND1':
            self.fields['CourseCode'].choices = Choices.HND1SecondSemester()
        elif Session == 'First Semester' and Class == 'HND2':
            self.fields['CourseCode'].choices = Choices.HND2FirstSemester()
        elif Session == 'Second Semester' and Class == 'HND2':
            self.fields['CourseCode'].choices = Choices.HND2SecondSemester()
        elif Session == 'Invalid' and Class == 'Invalid':
            self.fields['CourseCode'].choices = ()

        else:
            self.fields['CourseCode'].choices = Choices.allCourses()

class HodViewForm(forms.Form):
    Class = forms.ChoiceField(widget=forms.Select({'class':'selectField'}),choices=Choices.ClassChoices())
    Semester = forms.ChoiceField(widget=forms.Select({'class':'selectField'}),choices=Choices.SemesterChoice())
    Category = forms.ChoiceField(widget=forms.Select({'class':'selectField'}),choices=Choices.PendingOrAcceptedChoice())

class ExamOfficerViewForm(forms.Form):
    Class = forms.ChoiceField(widget=forms.Select({'class': 'selectField'}), choices=Choices.ClassChoices())
    Semester = forms.ChoiceField(widget=forms.Select({'class': 'selectField'}), choices=Choices.SemesterChoice())
    Category = forms.ChoiceField(widget=forms.Select({'class': 'selectField'}), choices=Choices.PendingOrCompiledChoice())

class chatForm(forms.Form):
    message = forms.CharField(widget=forms.Textarea(attrs={'class':'messageTextField','placehoder':'Message here'}), max_length=3000)

