<h1 align="center">
  <p> pycep-correios </p>
</h1>

<p align="center">
  API para busca de CEP integrado ao serviços dos Correios, ViaCEP e ApiCEP
</p>

## About 
  - É realizada a consulta ao serviço ViaCEP, sendo o CEP uma string podendo ou não ter pontuação.
  - Os CEP's a serem consultados estão dispostos em um arquivo .xlsx, não havendo interação com o usuário.
  - Após a consulta é retornado um dicionário, e as informações são escritas no arquivo .xlsx

## Return

```python
{
    'bairro': 'str',
    'cep': 'str',
    'cidade': 'str',
    'logradouro': 'str',
    'uf': 'str',
    'complemento': 'str',
}
```
## Install PyCEPCorreios
```cmd
pip install pycep-correios
```
## How to run program
  ```cmd
  python main.py
  ```
  
