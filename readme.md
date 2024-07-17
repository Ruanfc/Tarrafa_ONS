# Tarrafa
Tarrafa é uma aplicação em Python que serve para identificar arquivos que possuem um padrão Regex. Funciona como o comando *grep* do Linux, porém com prévia conversão de arquivos docx e pdf para txt. O diferencial da ideia é a praticidade da interface, ainda em ambiente Windows.

## Instalação
É recomendada a criação de um venv do python para evitar que exista conflito de versionamento entre as dependências dos pacotes instalados no ambiente base do pip. No terminal, crie o ambiente venv e o ative. Em seguida, clone o projeto em uma pasta de sua escolha e baixe os pacotes necessários por meio de pip install -r requirements.txt.

## Uso direto
A partir disso, o código pode ser rodado por python .\tarrafa_frontend.py.

## Criação de executável
Após a instalação, o executável pode ser criado a partir do pyinstaller