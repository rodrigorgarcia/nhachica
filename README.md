# Caminho de Nhá Chica

App de bolso para a peregrinação de 20/11 a 01/12/2026: roteiro dia a dia (km, terreno, pouso e telefones), passagens de ida e volta com QR code, contatos e informações gerais.

Tudo fica em um único arquivo, `index.html`, que funciona sem internet.

## Como usar no celular

- **Pelo GitHub Pages:** em *Settings → Pages*, escolha publicar a partir do branch com o `index.html`. Abra o endereço no celular uma vez com internet e use *Adicionar à tela inicial*. Depois disso o app abre mesmo sem sinal.
- **Como arquivo:** mande o `index.html` para o celular e abra no navegador. Tudo funciona sem JavaScript, exceto o destaque do dia de hoje e a lista "O que levar" que lembra o que já foi marcado.

## Como atualizar

Edite `Caminho_Nha_Chica_Roteiro.xlsx` e rode:

```
pip install openpyxl pypdf pillow
python build.py
```

O script lê as abas *Roteiro* e *Informações Gerais*, ignora a coluna K (Hospedagem Sugerida) e extrai os QR codes dos PDFs das passagens. Os demais dados das passagens ficam em `PASSAGENS`, no topo do `build.py`. O visual está em `src/app.html`.
