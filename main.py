from tqdm import tqdm
import os

from cep import Cep

os.system('cls')
cep = Cep()

for line in tqdm (range(cep.first_line, cep.last_line)):
    cep.read_data(line)
    cep.search_cep()
    cep.insert_data(line)

print('\n', '_' * 50)
input(' Finalizado. Aperte <ENTER> para fechar o programa\n')
