
import xml.etree.ElementTree as ET
import sys
from sys import argv

def parseToXml(arg):
    global root
    print("Reading " + arg)
    tree = ET.parse(arg)
    root = tree.getroot()