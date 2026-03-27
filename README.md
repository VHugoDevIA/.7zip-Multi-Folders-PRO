# ZIP Múltiplas Pastas - PRO v6.6

Aplicação **PowerShell** com interface gráfica Windows Forms para automatizar a compressão de múltiplas pastas usando o motor do 7-Zip (7z.exe).[file:1]  
Focada em cenários de produção, backups e preparação de ficheiros para envio, permite configurar parâmetros avançados de compressão, encriptação e gestão automática de logs.[file:1]

---

## Funcionalidades principais

- Interface gráfica (WinForms) para seleção e gestão de várias pastas numa só operação.[file:1]  
- Suporte para formatos `zip` e `7z`, com controlo detalhado de nível e método de compressão.[file:1]  
- Definição do número de threads de CPU a utilizar, respeitando o número de cores da máquina.[file:1]  
- Geração automática de nomes de arquivo com sufixo baseado no nome do computador (por exemplo, `PASTA_PC01.zip`).[file:1]  
- Estratégias configuráveis quando o ficheiro de saída já existe:
  - Criar sempre um novo arquivo,
  - Ignorar se já existir,
  - Gerar nome único incremental (`_2`, `_3`, ...).[file:1]
- Possibilidade de passar caminho **relativo** ou **absoluto** ao 7-Zip (controlo do layout interno do arquivo).[file:1]  
- Suporte a volumes divididos (multi-part), através de tamanho configurável (ex.: `100m`, `1g`).[file:1]  
- Opção para compressão de ficheiros partilhados (locked/shared).[file:1]  
- Criação opcional de executáveis SFX (`.exe`) quando se usa formato 7z.[file:1]  
- Encriptação com palavra-passe, incluindo:
  - ZIP: ZipCrypto, AES128, AES192, AES256,
  - 7Z: AES256 com opção de encriptar também os nomes dos ficheiros (headers).[file:1]
- Campo para parâmetros extra a passar diretamente ao 7-Zip (para utilizadores avançados).[file:1]  
- Log detalhado por sessão, com:
  - Ficheiro de log criado na pasta “Documentos” do utilizador, em `.ZIP_Múltiplas_Pastas-PRO LOG`,
  - Registo de comandos enviados ao 7-Zip, percentagens de progresso, erros e avisos,
  - Opção para abrir automaticamente a pasta do log no fim.[file:1]  
- Gestão de cancelamento:
  - Botão de “Cancelar” com confirmação,
  - Kill controlado do processo 7z.exe,
  - Limpeza dos arquivos criados nesta operação, se o utilizador cancelar.[file:1]
- Modo “eliminação segura” dos originais:
  - Se a opção estiver ativa, as pastas de origem são removidas apenas se toda a operação terminar com sucesso.[file:1]  
- Persistência de estado entre sessões:
  - Guarda formato, métodos, parâmetros, opções e posição/tamanho da janela em ficheiro JSON no `APPDATA`.[file:1]

---

## Requisitos

- Windows com suporte a PowerShell (recomendado PowerShell 5.1 ou superior).[file:1]  
- 7-Zip instalado em `C:\Program Files\7-Zip\7z.exe`.[file:1]  
- Permissões para executar scripts PowerShell (ajustar `ExecutionPolicy` se necessário).[file:1]

---

## Instalação

1. Instalar o 7-Zip na localização por defeito (`C:\Program Files\7-Zip\`).[file:1]  
2. Copiar o script `ZIP_Multiplas_Pastas_PRO_v6_6.ps1` para uma pasta à escolha.[file:1]  
3. Garantir que a política de execução permite correr o script, por exemplo:
   ```powershell
   Set-ExecutionPolicy RemoteSigned
