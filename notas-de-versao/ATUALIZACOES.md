# Atualizações do Organizador de Vista

## v3.25.0

- **Atualização automática pelo GitHub.** É o mesmo sistema do Super Captura: o
  programa procura versões novas sozinho ao abrir, baixa o pacote, confere o
  SHA-256 e a assinatura de cada arquivo, instala numa transação com backup e
  reabre. Se algo falhar no meio, volta tudo ao que era. A guia **Atualização**,
  na lateral, mostra a versão instalada e tem os botões Verificar e Atualizar.
  Enquanto uma operação do Tekla estiver rodando, a instalação espera.
- **Seus arquivos não são tocados pela atualização.** O `config.json`, o
  `tekla-root.txt` e a pasta `relatorios` ficam de fora do pacote de propósito.
- **Novo modo de abertura, igual ao do Super Captura.** O
  `Iniciar Organizador de Vista.bat` abre o programa sem a janela preta do
  console. O `Adicionar ao Menu Iniciar.bat` cria o atalho no Menu Iniciar, com
  ícone próprio, apontando direto para o `pyw.exe`.
- **Ícone próprio**, desenhado a partir do logo que já estava no cabeçalho do
  programa. Aparece na barra de tarefas, na janela e no Menu Iniciar.
- O `Abrir Organizador de Vista.bat` continua existindo, agora só para quando
  você quiser ver os logs e erros no console.

## v3.24 e anteriores

Ver os arquivos `ALTERACOES - v.3.22.txt`, `ALTERACOES - v.3.23.txt` e
`ALTERACOES - v.3.24.txt` na pasta do programa.
