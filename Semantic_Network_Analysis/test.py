#-*-coding: utf-8 -*-
import os

def add(a, b):
    return a+b;

def writeToTextFile(a):
    fileWriter=open("test.txt",'w')
    fileWriter.write(str(a)+" is writed")

def readParameters():
    # file=open("properties.txt")
    # parameters = file.read().split(";")
    file = '3;5'
    parameters = file.split(";")
    return parameters

def main():
    parameters = readParameters()
    a = int(parameters[0])
    b = int(parameters[1])
    writeToTextFile(add(a,b))
    
if __name__=="__main__":
    main()