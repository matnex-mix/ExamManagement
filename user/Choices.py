
def ClassChoices():
        Hnd1 = 'HND1'
        Hnd2 = 'HND2'
        Nd1Ft = 'ND1 FT'
        Nd2Ft = 'ND2 FT'
        NdYr1 = 'ND YR 1'
        NdYr2 = 'ND YR 2'
        Select = ''
        Choice = [
            (Select, 'Select A class'),
            (Hnd1, 'HND1'),
            (Hnd2, 'HND2'),
            (Nd1Ft, 'ND1 FT'),
            (Nd2Ft, 'ND2 FT'),
            (NdYr1, 'ND YR 1'),
            (NdYr2, 'ND YR 2')
        ]
        return Choice

def SessionChoice():
    Session2019_2020 = '2019 / 2020'
    Session2020_2021 = '2020 / 2021'
    Session2021_2022 = '2021 / 2022'
    Session2022_2023 = '2022 / 2023'
    SelectSession = ''
    Choice = [
        (SelectSession, 'Select a Session'),
        (Session2019_2020, '2019 / 2020'),
        (Session2020_2021, '2020 / 2021'),
        (Session2021_2022, '2021 / 2022'),
        (Session2022_2023, '2022 / 2023')
    ]
    return Choice

def SemesterChoice():
    FirstSemester = 'First Semester'
    SecondSemester = 'Second Semester'
    SelectSemester = ''
    Choice = [
        (SelectSemester, 'Select A semester'),
        (FirstSemester, 'First Semester'),
        (SecondSemester, 'Second Semester')
    ]
    return Choice

def UnitChoice():
    SelectUnit = ''
    TwoUnit = '2'
    ThreeUnit = '3'
    FourUnit = '4'
    Choice = [
        (SelectUnit, 'Select A Unit'),
        (TwoUnit, '2'),
        (ThreeUnit, '3'),
        (FourUnit, '4')
    ]
    return Choice

def HND1FirstSemester():
    Select = ''
    Com311 = 'COM 311'
    Com312 = 'COM 312'
    Com313 = 'COM 313'
    Sta311 = 'STA 311'
    Choices = [
        (Select, 'Select a Course'),
        (Com311, 'COM 311'),
        (Com312, 'COM 312'),
        (Com313, 'COM 313'),
        (Sta311, 'STA 311'),
    ]
    return Choices

def HND1SecondSemester():
    Select = ''
    Com321 = 'COM 321'
    Com322 = 'COM 322'
    Com323 = 'COM 323'
    Sta321 = 'STA 321'
    Choices = [
        (Select, 'Select a Course'),
        (Com321, 'COM 321'),
        (Com322, 'COM 322'),
        (Com323, 'COM 323'),
        (Sta321, 'STA 321'),
    ]
    return Choices

def HND2SecondSemester():
    Select = ''
    Com421 = 'COM 421'
    Com422 = 'COM 422'
    Com423 = 'COM 423'
    Com424 = 'Com 424'
    Choices = [
        (Select , 'Select a course'),
        (Com421 , 'COM 421'),
        (Com422 , 'COM 422'),
        (Com423 , 'COM 423'),
        (Com424 , 'Com 424')
    ]
    return Choices

def HND2FirstSemester():
    Select = ''
    Com411 = 'COM 411'
    Com412 = 'COM 412'
    Com413 = 'COM 413'
    Com414 = 'Com 414'
    Com415 = 'COM 415'
    Com416 = 'COM 416'
    Com417 = 'COM 417'
    Sta411 = 'Sta 411'
    Eed413 = 'EED 413'
    Gns401 = 'GNS 401'
    Gns412 = 'GNS 412'

    Choices = [
        (Select , 'Select a course'),
        (Com411 , 'COM 411'),
        (Com412 , 'COM 412'),
        (Com413 , 'COM 413'),
        (Com414 , 'Com 414'),
        (Com415 , 'COM 415'),
        (Com416 , 'COM 416'),
        (Com417 , 'COM 417'),
        (Sta411 , 'Sta 411'),
        (Eed413 , 'EED 413'),
        (Gns401 , 'GNS 401'),
        (Gns412 , 'GNS 412')
        ]
    return Choices

def allCourses():
    all = HND2FirstSemester() + HND1FirstSemester() + HND1SecondSemester() + HND2SecondSemester()
    return all

def PendingOrAcceptedChoice():
    Select = ''
    Pending = 'Pending'
    Accepted = 'Accepted'
    Choice =[
        (Select , 'Select a choice'),
        (Pending , 'Pending'),
        (Accepted , 'Accepted')
        ]
    return Choice

def PendingOrCompiledChoice():
    Select = ''
    Pending = 'Pending'
    Accepted = 'Compiled'
    Choice =[
        (Select , 'Select a choice'),
        (Pending , 'Pending'),
        (Accepted , 'Compiled')
        ]
    return Choice
