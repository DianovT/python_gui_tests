import pytest
from comtypes.client import *
import os

from fixture.application import Application


fixture = None


@pytest.fixture(scope="session")
def app(request):
    fixture = Application("C:\\Users\\dianov\\Downloads\\FreeAddressBookPortable\\AddressBook.exe")
    request.addfinalizer(fixture.destroy)
    return fixture


def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("xlsx_"):
            testdate = load_from_xslx(fixture[5:])
            metafunc.parametrize(fixture, testdate)

def load_from_xslx(file):
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), ("generator/%s.xlsx" % file))
    app = CreateObject("Excel.Application")
    workbooks = app.Workbooks.Open(file)
    worksheet = workbooks.Sheets[1]
    a = []
    for row in range(1, 100):
        data = worksheet.Cells[row, 1].Value()
        if data is not None:
            a.append(data)
    return a

