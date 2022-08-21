from random import randrange
from re import compile, search


class Generator:
    def gen_code_v2(self, x):
        """
        
        Parameters
        ----------
        x : int
            x represents quantity of chars that code is built from
    
        Returns
        -------
        code_list : list
            list of randomized codes that pass the regex
    
        """
        code_list = []
        reg = "^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[@$!%*#?&])[A-Za-z\d@$!#%*?&]{6,20}$"
        for times in range(10):
            code_generate = set(chr(randrange(33, 127)) for chars in range(11))
            code_list.append(''.join(code_generate))
        for code in code_list:
                comp = compile(reg)
                check = search(comp, code)
                if check:
                    pass
                else:
                    code_list.remove(code)
        return code_list

code = Generator()
x = int(input('How many chars you want your code to be? (min. 6): '))
if x < 6:
    print('Minimum is 6!')
else:
    print(f'Here are your generated codes: {code.gen_code_v2(x)}')