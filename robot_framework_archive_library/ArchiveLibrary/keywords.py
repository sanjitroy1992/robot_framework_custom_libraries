#!/usr/bin/env python

from robot.libraries.Collections import Collections
from robot.libraries.OperatingSystem import OperatingSystem
from ArchiveLibrary.utils import Unzip, Untar, return_files_lists
import os
import tarfile
import zipfile
import subprocess
import shutil
# from Libraries.Common.GenericLib import GenericLib
import subprocess
from subprocess import Popen, PIPE

class ArchiveKeywords(object):
    ROBOT_LIBRARY_SCOPE = 'Global'

    tars = ['.tar', '.tar.bz2', '.tar.gz', '.tgz', '.tz2']

    zips = ['.docx', '.egg', '.jar', '.odg', '.odp', '.ods', '.xlsx', '.odt',
            '.pptx', 'zip']

    def __init__(self):
        self.oslib = OperatingSystem()
        self.collections = Collections()

    def _killAllProcess(self, *processList):
        """ Used to Kill all running process by passing list of processname.

         Arguments: '*processList' contains variable number of processname

         Example:

        | ***Variable*** |
        | @{AllProcessToKill} | chrome.exe | chromedriver.exe |
        | ***TestCases*** |
        | KillAllProcess | @{AllProcessToKill} |

        """
        tasklist = subprocess.Popen(("tasklist"), stdout=PIPE, stderr=PIPE)
        stdout, stderr = tasklist.communicate()
        for process in processList:
            if process in str(stdout):
                killCommand = "Taskkill /IM" + " " + process + " /F"
                try:
                    subprocess.Popen(killCommand, stdout=PIPE, stderr=PIPE)
                except:
                    pass
        filesintemp = [name for name in os.listdir(os.environ['TEMP'])]
        for i in filesintemp:
            if "scoped_dir" in i:
                dir = os.path.join(os.environ['TEMP'], i)
                try:
                    shutil.rmtree(dir)
                except:
                    pass

    def _extract_zip_file(self, zfile, dest=None):
        ''' Extract a ZIP file

        `zfile` the path to the ZIP file

        `dest` optional destination folder. Assumes current working directory if it is none
               It will be created if It doesn't exist.
        '''

        if dest:
            self.oslib.create_directory(dest)
            self.oslib.directory_should_exist(dest)
        else:
            dest = os.getcwd()

        cwd = os.getcwd()

        unzipper = Unzip()

        # Dont know why I a have `gotta catch em all` exception handler here
        try:
            unzipper.extract(zfile, dest)
        except:
            raise
        finally:
            os.chdir(cwd)

    def _extract_tar_file(self, tfile, dest=None):
        ''' Extract a TAR file

        `tfile` the path to the TAR file

        `dest` optional destination folder. Assumes current working directory if it is none
               It will be created if It doesn't exist.
        '''
        if dest:
            self.oslib.create_directory(dest)
        else:
            dest = os.getcwd()

        self.oslib.file_should_exist(tfile)

        untarrer = Untar()
        untarrer.extract(tfile, dest)

    def archive_should_contain_file(self, zfile, filename):
        ''' Check if a file exists in the ZIP file without extracting it

        `zfile` the path to the ZIP file

        `filename` name of the file to search for in `zfile`
        '''
        self.oslib.file_should_exist(zfile)

        files = []
        if zipfile.is_zipfile(zfile):
            files = zipfile.ZipFile(zfile).namelist()
        else:
            files = tarfile.open(name=zfile).getnames()
        files = [os.path.normpath(item) for item in files]

        self.collections.list_should_contain_value(files, filename)

    def _create_tar_from_files_in_directory(self, directory, filename, sub_directories=True):
        """ Take all files in a directory and create a tar package from them

        `directory` Path to the directory that holds our files

        `filename` Path to our destination TAR package.

        `sub_directories` Shall files in sub-directories be included - True by default.
        """
        tar = tarfile.open(filename, "w")
        files = return_files_lists(directory, sub_directories)

        for filepath, name in files:
            tar.add(filepath, arcname=name)

        tar.close()

    def _create_zip_from_files_in_directory(self, directory, filename, sub_directories=False):
        """ Take all files in a directory and create a zip package from them

        `directory` Path to the directory that holds our files

        `filename` Path to our destination ZIP package.

        `sub_directories` Shall files in sub-directories be included - False by default.
        """
        the_zip = zipfile.ZipFile(filename, "w")
        files = return_files_lists(directory, sub_directories)

        for filepath, name in files:
            the_zip.write(filepath, arcname=name)

        the_zip.close()

    def extract_file(self, compressed_file_path, destination=None):
        """|Usage|
            Extract the contents of the compressed 'file' into a 'dest' path.

            |Arguments|
            file: Path to the package (zip, tar, war, jar, ear).
            dest: Destination path to extract the package. By default takes current directory.

            |Example|
            *** Variables ***
            ${destination}    D:/extracted_data
            *** Test Cases ***
            Extract File    C:/Uxp.war    ${destination}
            """
        if not os.path.exists(compressed_file_path):
            raise AssertionError("File not found error : " + compressed_file_path + " doesnot exist")
        head, tail = os.path.split(compressed_file_path)
        ext = tail.split(".")[1]
        if ext.lower() == 'zip':
            self._extract_zip_file(compressed_file_path, destination)
        elif ext.lower() == 'tar':
            self._extract_tar_file(compressed_file_path, destination)
        elif ext.lower() in ['war', 'jar', 'ear']:
            processList = ['7zFM.exe', 'jar.exe']
            self._killAllProcess(*processList)
            if os.path.exists(destination):
                shutil.rmtree(destination)
            os.mkdir(destination)
            if not destination:
                a = subprocess.Popen("jar xf {}".format(compressed_file_path), cwd=os.curdir)  #Extracts the data in current project directory
                a.wait()
            else:
                result = subprocess.Popen("jar xf {}".format(compressed_file_path), cwd=destination, stdout=subprocess.PIPE)    #Extracts the data in destination directory specified
                result.wait()
                return result
        else:
            raise AssertionError("Invalid format : " + tail)

    def create_compressed_file_from_files_in_directory(self, directory, compressed_file_path, sub_directories=False):
        """|Usage|
            Take all files specified in the 'directory' and create a compressed package from them and store it in the specified 'compressed_file_path'.

            |Arguments|
            directory: Path to the directory that holds our files.
            compressed_file_path: Path to store our destination package (zip, tar, war, jar, ear).
            sub_directories: Shall files in sub-directories be included - False by default.

            |Example|
            *** Variables ***
            ${directory}    D:/extractedjardata
            *** Test Cases ***
            Create Compressed File From Files In Directory    ${directory}   C:/BOHistoryDetails.jar
        """
        head, tail = os.path.split(compressed_file_path)
        if not os.path.exists(head):
            raise AssertionError("Directory not found error : " + head + " doesnot exist")
        if not os.path.exists(directory):
            raise AssertionError("Directory not found error : " + directory + " doesnot exist")
        processList = ['7zFM.exe', 'jar.exe']
        self._killAllProcess(*processList)
        if os.path.exists(compressed_file_path):
            os.remove(compressed_file_path)
        ext = tail.split(".")[1]
        if ext.lower() == 'zip':
            self._create_zip_from_files_in_directory(directory, compressed_file_path, sub_directories)
        elif ext.lower() == 'tar':
            self._create_tar_from_files_in_directory(directory, compressed_file_path, sub_directories)
        elif ext.lower() in ['war', 'jar', 'ear']:
            abc = os.listdir(directory)
            a = ""
            for i in abc:
                a = a + " " + i
            result = subprocess.Popen("jar cf {} {}".format(compressed_file_path, a), cwd=directory)
            result.wait()
            return result
        else:
            raise AssertionError("Invalid format : " + tail)


if __name__ == '__main__':
    al = ArchiveKeywords()
    al.extract('test.zip')
