# Atualizações do Organizador de Vista

## v3.27.0

- **Acabou o flash da janela preta ao abrir.** O programa agora é o
  **`Organizador de Vista.pyw`** — dois cliques nele e pronto. Arquivo `.pyw`
  abre direto no `pyw.exe` do Python, que não tem console nenhum.
- O flash vinha do `.bat`: todo arquivo `.bat` abre uma janela do `cmd` por um
  instante, mesmo quando o comando dentro dele não mostra nada. Por isso o
  `Abrir Organizador de Vista.bat` foi removido — a atualização apaga ele
  sozinha de quem já tinha.
- O `Adicionar ao Menu Iniciar.bat` continua, apontando para o `.pyw`. Use ele
  uma vez e depois abra pelo Menu Iniciar, com ícone e tudo.

## v3.26.1

- **Abrir o programa não interrompe mais o trabalho.** A verificação de
  atualização continua acontecendo sozinha ao abrir, mas agora ela só deixa a
  guia Atualização pronta, sem abrir nada na frente. Quando quiser, você vai lá
  e clica em Atualizar.
- A única exceção: se a mesma versão nova ficar **mais de 7 dias** sem ser
  instalada, aí sim o programa avisa na abertura. Versão mais nova reinicia a
  contagem.
- **Organizar, Otimizar e Desfazer passaram para dentro do cartão do Plano de
  Organização**, separados do resultado por uma linha fina. O cartão termina
  exatamente na mesma altura da barra lateral.

## v3.26.0

- **O Registro de Atividades saiu da tela.** A tabela do Plano de Organização
  ocupa agora a altura toda da janela.
- **No lugar dele, notificação no centro da janela**, no estilo do programa:
  uma por operação, com o resultado em uma frase direta. Verde para concluído,
  amarelo para atenção, vermelho para falha.
- **Nada mais falha em silêncio.** Antes, se algo desse errado, a mensagem ia
  só para aquele registro no rodapé — e a caixa de aviso da interface estava
  desativada por um `display: none`, então alertas não apareciam em lugar
  nenhum. Agora toda falha e todo aviso abrem a notificação.
- **As linhas do processo continuam disponíveis** dentro da notificação, em
  "Detalhes técnicos", fechado. Quem só quer o resultado lê a frase e fecha;
  quem precisa investigar abre.

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
