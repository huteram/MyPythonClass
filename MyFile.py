# -*- coding: cp1250 -*-
#version 2021 - 03
#created by: Milan Hutera
#class for working with folders and files
import os
import re
import fnmatch
import mmap
from glob import glob


class FileSystem:
    def __addFolderSpliter__(self, Folder, fullPath = True):       
        """ adjust the folder name for other functions """
        def reduceSpecialChar(Directory):
            ListDir = Directory.split("\\")
            while ListDir.count("..") > 0 :
                Index = ListDir.index("..")
                Value = ListDir.pop(Index)
                Value = ListDir.pop(Index-1)
            result = "\\".join(map(str,ListDir))
            return result




        if fullPath:     
            if Folder.count(":")  == 0:
                resultFolder = os.getcwd() + '\\'
            else:
                resultFolder = ""
        else:
            resultFolder = ""        
        resultFolder += Folder

        if len(resultFolder)>0:
            if resultFolder[-1] != "\\":
                resultFolder += "\\"
        else:      
            resultFolder = (os.getcwd() + Folder)

        resultFolder = reduceSpecialChar(resultFolder)
        return resultFolder


    def listFile(self, Folder, Pattern = "*.*"):
        """ return list of files in folder and subfolders filtered by pattern """
        result = list()
        Folder = self.__addFolderSpliter__(Folder)

        for dirPath, dirNames, fileNames in os.walk(Folder):
            for fileName in fileNames:
                if fnmatch.fnmatch(fileName, Pattern):
                    result.append(self.__addFolderSpliter__(dirPath) + fileName)
        return result
    
    def localDirectory(self, currentFile = __file__):
        BaseDirectory = currentFile[:currentFile.rfind('\\')+1]
        return BaseDirectory
        
    def deleteFiles(self, Folder, Pattern='*.*', PrintListFiles = True):
        """Delete all files in folder and subfolders filtered by pattern """
        deletedFiles = self.listFile(Folder, Pattern)      
        for File in deletedFiles:
            os.remove(File)  
            if PrintListFiles:
                print("deleting: ", File)      
        return deletedFiles
    
    def moveFiles(self, Folder, NewFolder, Pattern='*.*', PrintListFiles = True):
        """Delete all files in folder and subfolders filtered by pattern """
        movedFiles = self.listFile(Folder, Pattern)
        NewFolder = self.__addFolderSpliter__(NewFolder)
        Folder = self.__addFolderSpliter__(Folder)
        for File in movedFiles:
            FileName = File.split('\\')[-1]
            subFolderFile = File.split(Folder)[-1]
            self.MakeDirectory(NewFolder + subFolderFile.split(FileName)[0])
            os.rename(File,NewFolder + subFolderFile)
            if PrintListFiles:
                print("moved: ", File)      
        return movedFiles
    
    def deleteDirs(self, Directory):
        """ delete all subfolders in directory and return the list of deleting dirs"""

        deletedDirs = list()
        Directory = self.__addFolderSpliter__(Directory)

        # walk through the directories and make a list of dirs. 
        for dirpath, dirnames, filenames in os.walk(Directory):
            for DirName in dirnames:
                DeleteName = os.path.join(dirpath, DirName)
                deletedDirs.append(DeleteName)
        deletedDirs.reverse()

        # delete all dirs in the list. 
        for DeleteItem in deletedDirs:
            print('Deleting directory: ', DeleteItem)

            # if directory is not empty delete all the files
            ListDir = os.listdir(DeleteItem)
            if (len(ListDir)>0):
                self.deleteFiles(DeleteItem, "*")

            if (os.path.isdir(DeleteItem)):
                os.removedirs(DeleteItem)

        return deletedDirs

    def MakeDirectory(self, Directory):
        """ Make a new dirctory and subdirectory if it is necessary """
        Directory = self.__addFolderSpliter__(Directory)
        result = False 
        # if dirctory already exists the result keeps False. 
        DirectoryWay = Directory.split('\\')
        CurrentDirectory = DirectoryWay[0]
        for directoryItem in DirectoryWay[1:]:
            CurrentDirectory += "\\"
            CurrentDirectory += directoryItem
            if not(os.path.isdir(CurrentDirectory)):
                os.mkdir(CurrentDirectory)
                result = True        
        return result
    
    def FileExists(self,File):
        "Return true if the file exists."
        exists = os.path.isfile(File)
        return exists

    def readFile(self, FileName):
        if self.FileExists(FileName):
            try:
                File = open(FileName,'r+')
                FileLength = os.path.getsize(FileName)
                DataFile = mmap.mmap(File.fileno(),FileLength)
                Data = DataFile.read(FileLength)
                DataFile.close()
            except Exception as Error:               
                Data = -1
                print ('reading file "{}" failed: {}'.format(FileName, Error))
        else:
            print('file "{}" doesn`t exist.'.format(FileName))
            Data = -1        
        return Data

    def ListDir(self, Adresar=''):
        if Adresar =='' or Adresar[0]== '\\':
            Adresar = (os.getcwd() + Adresar)
            
        if Adresar.find('*')<0:
            SeznamAdresaru = [s for s in os.listdir(Adresar) if os.path.isdir(s)]
        else:
            SeznamAdresaru = [s for s in glob(Adresar) if os.path.isdir(s)]
            s = []
            for i in SeznamAdresaru:
                s.append(i[i.rfind('\\')+1:])
            SeznamAdresaru = s
        return SeznamAdresaru

    def rename(self, FileName, NewName):
        if self.FileExists(FileName):
            try:
                NewName = NewName.split('\\')[-1]
                FileNewName = ""
                
                for i in FileName.split('\\')[:-1]:
                    FileNewName += i
                    FileNewName += "\\"
                FileNewName += NewName
                os.rename(FileName, FileNewName)
            except Exception as Error:               
                Data = -1
                print ('renaming file "{}" failed: {}'.format(FileName, Error))
        else:
            print('file "{}" doesn`t exist.'.format(FileName))
            
            
        
            



if __name__ == '__main__':
    testClass = FileSystem()
    # List = testClass.deleteFiles("..\\testFolder","*.py")
    # ListDir = testClass.deleteDirs("..\\testFolder")
    c = os.getcwd()
    
    # testClass = "..\\testFolder\\MujSoubor.pyc"
    s = "..\\testFolder\\MujSoubor.pyc"




    




