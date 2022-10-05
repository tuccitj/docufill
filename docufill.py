from docx import Document
import os
import json

# set current_directory_path to current directory path

current_directory_path = os.path.dirname(__file__)
input_path = current_directory_path + '/dictionary'
templates_path = current_directory_path + '/templates'
output_path = current_directory_path + '/output'

# Checks if required folders exist, if not, creates them


def initialize_directory():
    # init list of paths to create
    paths_to_create = [
        current_directory_path + '/dictionary',
        current_directory_path + '/templates',
        current_directory_path + '/output'
    ]
    # for each path in list, create directory, if it exists, ignore error
    for path in paths_to_create:
        os.makedirs(path, exist_ok=True)


def get_lists_of_files():
    # files = list()
    # files.append(os.listdir(input_path))
    # files.append(os.listdir(templates_path))
    files = {
        'inputs': os.listdir(input_path),
        'templates': os.listdir(templates_path)
    }
    return files


def importDictionary(inputs):
    input_file_path = input_path + "/dictionary.txt"
    with open(input_file_path) as file:
        dictionary = file.read()
        # print(dictionary)
        dict = json.loads(dictionary)
        return dict
    # return dictionary


def findAndReplace(templates, dictionary):
    for file in templates:
        file_path = templates_path + "/" + file
        document = Document(file_path)
        for p in document.paragraphs:
            inline = p.runs
            for i in range(len(inline)):
                text = inline[i].text
                for key in dictionary.keys():
                    if key in text:
                        text = text.replace(key, dictionary[key])
                        inline[i].text = text

        file_output_path = output_path + "/docufilled-" + file
        document.save(file_output_path)
        output = {'input_path': file_path,
                  'output_path': file_output_path}
        print(output)


def main():
    print("start init")
    initialize_directory()
    print("finished init")
    print("start get lists of files")
    files = get_lists_of_files()
    print("finished get lists of files")
    print("start get dict")
    dictionary = importDictionary(files['inputs'])
    print("finished get dict")
    #print("dictionary: ", dictionary)
    print("start far")
    findAndReplace(files['templates'], dictionary)
    print("finish far")
    print("Finished.")


main()
