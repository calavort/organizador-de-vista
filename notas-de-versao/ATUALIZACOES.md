# Atualizações do Organizador de Vista

## v3.25.3

- **Tabela do Plano de Organização com as colunas redistribuídas.** Posição,
  Vista e Escala sobravam espaço; Status faltava, e "Planejada - grupo
  empilhado" aparecia cortado. O espaço que sobrava foi para onde faltava: o
  texto aparece inteiro, mesmo com a janela no tamanho mínimo.

## v3.25.2

- Um único atalho para abrir: o `Abrir Organizador de Vista.bat`, agora sem a
  janela preta do console. O `Iniciar Organizador de Vista.bat` era repetição e
  foi removido — a atualização apaga ele sozinha de quem já tinha os dois.
- A guia **Atualização** não cabia na faixa e aparecia com o nome cortado.
- Guia mais limpa: saíram a nota explicativa e o crédito repetido (o rodapé do
  programa já assina).

## v3.25.1

- O instalador agora espera o programa fechar de verdade antes de trocar os
  arquivos, e nenhuma segunda janela consegue abrir no meio da instalação.

## v3.25.0

- **Atualização automática pelo GitHub.** É o mesmo sistema do Super Captura: o
  programa procura versões novas sozinho ao abrir, baixa o pacote, confere o
  SHA-256 e a assinatura de cada arquivo, instala numa transação com backup e
  reabre. Se algo falhar no meio, volta tudo ao que era. A guia **Atualização**,
  na lateral, mostra a versão instalada e tem os botões Verificar e Atualizar.
  Enquanto uma operação do Tekla estiver rodando, a instalação espera.
- **Seus arquivos não são tocados pela atualização.** O `config.json`, o
  `tekla-root.txt` e a pasta `relatorios` ficam de fora do pacote de propósito.
- **Novo modo de abertura, igual ao do Super Captura.** O programa abre sem a
  janela preta do console, e o `Adicionar ao Menu Iniciar.bat` cria o atalho no
  Menu Iniciar, com ícone próprio, apontando direto para o `pyw.exe`.
- **Ícone próprio**, desenhado a partir do logo que já estava no cabeçalho do
  programa. Aparece na barra de tarefas, na janela e no Menu Iniciar.

## v3.24 e anteriores

Ver os arquivos `ALTERACOES - v.3.22.txt`, `ALTERACOES - v.3.23.txt` e
`ALTERACOES - v.3.24.txt` na pasta do programa.
