import pytest
import importlib
import os

from fixture.application import Application


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

def load_from_module(module):
    return importlib.import_module("data.%s" % module).testdate

def load_from_xslx(file):
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), ("generator/%s.xlsx" % file))) as f:
        return f.read()
