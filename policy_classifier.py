from doc import Doc
import sys, os 
from argparse import ArgumentParser 
from shutil import copy
from concurrent.futures import ProcessPoolExecutor

def main():
    dir = check_directory() #verify that the provided directory exists, if not then default to the current directory.
    print("Loading and analyzing files....")
    file_list = extract_files(dir[0]) # for each of the files in the provided directory, create file objects and add to list
    domains = build_domains(dir[1]) # create the list of domains from the default 'domains' folder, can add specific folder as argument
    score_file(file_list, domains) # score the files based on how well they match keywords in each of the domains
    rank_place(file_list, domains) # copy files to the domain-named folders they best match
    print("Complete. Check the \'Results\' folder for the sorted files.")


def check_directory():
    """
    Verify that provided path exists.
    """
    parser = ArgumentParser(description='Classify and sort a collection of PDF and docx files based on cybersecurity keywords.')
    parser.add_argument("-f", "--file", help="the folder to search and classify", required=False, default=os.getcwd())
    parser.add_argument("-d", "--directory", help="the folder containing domain txt files with keywords.", required=False, default="domains")
    path = parser.parse_args().file
    doms = parser.parse_args().directory
    
    if os.path.exists(path):
        return (path,doms)
    else: 
        sys.exit("Please enter in a valid file path.")

def extract_files(directory):
    """
    Read each of the files in the provided path, create a File Object for each of the files if its a pdf or docx. 
    """
    file_list = [] # list of file objects based on the file name from the walk
    for root, _, files in os.walk(directory):
        with ProcessPoolExecutor() as executor:
            future = [executor.submit(Doc, os.path.join(root, file)) for file in files if file.endswith(".docx") or file.endswith(".pdf")]
            for f in future:
                file_list.append(f.result())
    return file_list

def build_domains(file="domains"):
    """
    Gather all of the domains files based on the provided domains folder, and create a list with all the file names
    """
    return [domain.path for domain in os.scandir(file)]

def score_file(file_list, domains_list):
    """
        Score the Policy based on how well it matches the keywords from the provided domain.
    """
    for file in file_list:               
        for domain in domains_list: #write out as a scoring function so it can be called for pdf and word docs, currently does not work for pdf
            with open(domain) as dom:
                domtxt = [line.rstrip().lower() for line in dom]
            for word in domtxt:
                word_count = 0
                if word in file.text[0]: # full text
                    word_count = file.text[0].count(word)
                    if word in file.text[1]: # headers
                        word_count += file.text[1].count(word)*2
                    if word in file.text[2]: # bold / underline / italics
                        word_count += file.text[2].count(word)*3
                try: 
                    file.score[os.path.split(domain)[-1][:-4]] += word_count
                except KeyError:
                    file.add_domain(os.path.split(domain)[-1][:-4], word_count)


def rank_place(file_list, domain_list):
    """
    Create folder based on each of the domains, and then sort the files into the appropriate domains based on the keyword score in the file (set in score_file()) 
    """
    try:
        os.mkdir('Results')
        os.mkdir(os.path.join('.\Results', '0. all_policies'))
    except FileExistsError:
        #Folder already in place, ignore
        pass
    for domain in domain_list:
        try:
            os.mkdir(os.path.join("Results", os.path.split(domain)[-1][:-4]))
        except FileExistsError:
            #Folder already in place, ignore
            pass
    for file in file_list:
        for key in file.score:
            try: 
                if (file.score[key] / len(file.text[0].strip())) >= .003: 
                    copy(file.path, os.path.join('.\Results',key,file.name))
                else: 
                    copy(file.path, os.path.join('.\Results', '0. all_policies'))
            except ZeroDivisionError:
                print(f"Error reading {str(file)}, please manually review.")
                break

if __name__ == "__main__":
    main()
