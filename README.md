# Planilha Taticamente
Aplicativo feito em python para preencher automaticamente a planilha criada pelo site Taticamente.
Ao selecionar um ou mais arquivos html com os atributos dos jogadores, cada um terá a sua planilha criada com o mesmo 
nome do arquivo html.

[Veja aqui como utilizar a planilha](https://taticamente.com/atributos-dos-jogadores-no-football-manager/)

## Python
Quem tiver o python instalado, basta executar o aquivo main que está dentro da pasta src:
```commandline
python main.py
```

## Aplicativo
Também disponibilizei um aplicato .exe para que seja mais fácil de utilizar (para quem não tem conhecimentos de python).
Em caso de dúvidas deixei todos os arquivos que o python utiliza para criar um arquivo .exe no repositório.
E claro todo o código é open-source.

Pode ser feito o download na sesão "Releases" aqui ao lado.
O arquivo se chama "gerador_planilhas_taticamente.exe".

Como eu não tenho certificado para assinar o software, com certeza o windows vai perguntar se confia no executável e 
se pode executar assim mesmo.

Obs: Se a planilha do site Taticamente estiver na mesma pasta do arquivo .exe, ela será automaticamente carregada.

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

To-do:
- [ ] Assinar software
- [ ] diminuir o tamanho do executável
