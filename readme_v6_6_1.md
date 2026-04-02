# ZIP Múltiplas Pastas - PRO v6.6.1

**Alterações nesta versão (v6.6.1):**  
- Correções menores na v6.6, incluindo ajustes na interface e validações.

Aplicação **PowerShell** com interface gráfica Windows Forms para automatizar a compressão de múltiplas pastas usando o motor do 7-Zip (7z.exe).  
Focada em cenários de produção, backups e preparação de ficheiros para envio, permite configurar parâmetros avançados de compressão, encriptação e gestão automática de logs.

---

## Funcionalidades principais

- Interface gráfica (WinForms) para seleção e gestão de várias pastas numa só operação.  
- Suporte para formatos `zip` e `7z`, com controlo detalhado de nível e método de compressão.  
- Definição do número de threads de CPU a utilizar, respeitando o número de cores da máquina.  
- Geração automática de nomes de arquivo com sufixo baseado no nome do computador (por exemplo, `PASTA_PC01.zip`).  
- Estratégias configuráveis quando o ficheiro de saída já existe:
  - Criar sempre um novo arquivo,
  - Ignorar se já existir,
  - Gerar nome único incremental (`_2`, `_3`, ...).  
- Possibilidade de passar caminho **relativo** ou **absoluto** ao 7-Zip (controlo do layout interno do arquivo).  
- Suporte a volumes divididos (multi-part), através de tamanho configurável (ex.: `100m`, `1g`).  
- Opção para compressão de ficheiros partilhados (locked/shared).  
- Criação opcional de executáveis SFX (`.exe`) quando se usa formato 7z.  
- Encriptação com palavra-passe, incluindo:
  - ZIP: ZipCrypto, AES128, AES192, AES256,
  - 7Z: AES256 com opção de encriptar também os nomes dos ficheiros (headers).  
- Campo para parâmetros extra a passar diretamente ao 7-Zip (para utilizadores avançados).  
- Log detalhado por sessão, com:
  - Ficheiro de log criado na pasta "Documentos" do utilizador, em `.ZIP_Múltiplas_Pastas-PRO LOG`,
  - Registo de comandos enviados ao 7-Zip, percentagens de progresso, erros e avisos,
  - Opção para abrir automaticamente a pasta do log no fim.  
- Gestão de cancelamento:
  - Botão de "Cancelar" com confirmação,
  - Kill controlado do processo 7z.exe,
  - Limpeza dos arquivos criados nesta operação, se o utilizador cancelar.  
- Modo "eliminação segura" dos originais:
  - Se a opção estiver ativa, as pastas de origem são removidas apenas se toda a operação terminar com sucesso.  
- Persistência de estado entre sessões:
  - Guarda formato, métodos, parâmetros, opções e posição/tamanho da janela em ficheiro JSON no `APPDATA`.  

---

## Requisitos

- Windows com suporte a PowerShell (recomendado PowerShell 5.1 ou superior).  
- 7-Zip instalado em `C:\Program Files\7-Zip\7z.exe`.  
- Permissões para executar scripts PowerShell (ajustar `ExecutionPolicy` se necessário).  

---

## Instalação

1. Certifique-se de que o 7-Zip está instalado no caminho padrão.  
2. Execute o script `zip_multiplas_pastas_pro_v6_6_1.ps1` com PowerShell.  
3. As configurações são guardadas automaticamente no APPDATA.