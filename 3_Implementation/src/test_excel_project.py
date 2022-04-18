from .excel_project import SemesterMarks, HobbiesList, CitiesVisited, ProgrammingLanguage, SportsList

class TestSemesterMarks:
    def test_semester_fun1(self):
        s=SemesterMarks("Semester_Marks",99004400)
        assert s.semester_fun() == 1
    def test_semester_fun2(self):
        s=SemesterMarks("Semester_Marks",9900440)
        assert s.semester_fun()==0

class TestHobbiesList:
    def test_hobbies_fun1(self):
        h=HobbiesList("Hobbies_List",99004408)
        assert h.hobbies_fun() == 1
    def test_hobbies_fun2(self):
        h=HobbiesList("Hobbies_List",9900450)
        assert h.hobbies_fun() == 0

class TestCitiesVisited:
    def test_cities_fun1(self):
        c=CitiesVisited("Cities_Visited",99004301)
        assert c.cities_fun() == 0
    def test_cities_fun2(self):
        c=CitiesVisited("Cities_Visited",99004412)
        assert c.cities_fun() == 1

class TestProgrammingLanguage:
    def test_programming_fun1(self):
        p=ProgrammingLanguage("ProgrammingLanguage_Expertise",99004402)
        assert p.programming_fun() == 1
    def test_programming_fun2(self):
        p=ProgrammingLanguage("ProgrammingLanguage_Expertise",990002)
        assert p.programming_fun() == 0

class TestSportsList:
    def test_sports_fun1(self):
        s=SportsList("Sports_List",99004402)
        assert s.sports_fun() == 1
    def test_sports_fun2(self):
        s=SportsList("Sports_List",123456)
        assert s.sports_fun() == 0