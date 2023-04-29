class Printing:
    def __init__(self):
        pass

    class formatting:
        HEADER = '\033[95m'
        OKBLUE = '\033[94m'
        OKCYAN = '\033[96m'
        OKGREEN = '\033[92m'
        WARNING = '\033[93m'
        FAIL = '\033[91m'
        ENDC = '\033[0m'
        BOLD = '\033[1m'
        UNDERLINE = '\033[4m'

    def print_new(self):
        print("\n")

    def print_inline(self, str):
        print("    "+str, end="")
        self.backline()

    def backline(self):
        print('\r', end="")

    def welcome(self):
        print(f'''


    
    ███╗   ██╗███████╗██████╗ ██╗     
    ████╗  ██║██╔════╝██╔══██╗██║     
    ██╔██╗ ██║███████╗██████╔╝██║     
    ██║╚██╗██║╚════██║██╔══██╗██║     
    ██║ ╚████║███████║██████╔╝███████╗
    ╚═╝  ╚═══╝╚══════╝╚═════╝ ╚══════╝


    {self.formatting.HEADER}{self.formatting.OKBLUE}Starting to extract this weeks runsheet and scoresheets!{self.formatting.ENDC}



''')
