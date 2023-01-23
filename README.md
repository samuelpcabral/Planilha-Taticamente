# Planilha Taticamente
Aplicativo feito em pyhton para preencher automaticamente a planilha criada pelo site taticamente.
Ao selecionar um ou mais arquivos html com os atributos dos jogadores, cada um terá a sua planilha criada com o mesmo 
nome do arquivo html

[Veja aqui como utilizar a planilha](https://taticamente.com/atributos-dos-jogadores-no-football-manager/)

## Aplicativo
Também disponibilizei um aplicato .exe para que seja mais fácil de utilizar. (para quem não tem conhecimentos de python)
em caso de dúvida deixei todos os arquivos que o python utiliza para criar um arquivo exe no repositório.
E claro todo o código é open-source.

## Dev

### Pra quem quiser colaborar:

Package list:
```commandline
pip install pandas
pip install openpyxl
```

Para gerar o executável:
```commandline
pyinstaller --noconsole --icon=icon.ico -n gerador_planilhas_taticamente -F .\main.py
```
