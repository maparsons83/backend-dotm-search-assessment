#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""
import os
import sys
import zipfile

# Your awesome code begins here!
def dotm_search(text, dir):
    print "Searching directory {} for text {} ...".format(dir, text)
    filenames = os.listdir(dir)
    matches = 0
    searches = 0
    for filename in filenames:
        if filename.endswith('.dotm'):
            searches += 1
            try:
                zip_file = zipfile.ZipFile(os.path.abspath(os.path.join)(dir, filename), 'r')
                data = zip_file.read('word/document.xml')
                if data.find(text) != -1:
                    print 'Match found in file: {}'.format(os.path.join(dir, filename))
                    index = data.find(text)
                    print '...{}...\n'.format(data[(index - 40):(index + 41)])
            except:
                print "Problem with file: %s" % filename
                print "Total dotm files searched: {}".format(searches)
                print "Total dotm files matched: {}".format(matches)

if __name__ == '__main__':
    cwd = os.getcwd()
    if len(sys.argv) == 2:
        dotm_search(sys.argv[1], cwd)
    if len(sys.argv) == 4 and sys.argv[2] == '--dir':
        dotm_search(sys.argv[1], sys.argv[3])
    else:
        print 'usage: dotm_search.py <text> --dir "./dotm_files"'
        sys.exit(1)  


