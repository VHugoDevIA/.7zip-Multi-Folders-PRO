# ZIP Múltiplas Pastas - PRO v6.6

Aplicação **PowerShell** com interface gráfica Windows Forms para automatizar a compressão de múltiplas pastas usando o motor do 7-Zip (7z.exe).  
Focada em cenários de produção, backups e preparação de ficheiros para envio, permite configurar parâmetros avançados de compressão, encriptação e gestão automática de logs.

---

## Funcionalidades principais

- Interface gráfica (WinForms) para seleção e gestão de várias pastas numa só operação.  
- Suporte para formatos `zip` e `7z`, com controlo detalhado de nível e método de compressão.  
- Definição do número de threads de CPU a utilizar.  
- Geração automática de nomes de arquivo com sufixo baseado no nome do computador.  
- Estratégias configuráveis quando o ficheiro de saída já existe.  
- Possibilidade de passar caminho relativo ou absoluto ao 7-Zip.  
- Suporte a volumes divididos.  
- Opção para compressão de ficheiros partilhados.  
- Criação opcional de executáveis SFX.  
- Encriptação com palavra-passe (ZipCrypto, AES).  
- Campo para parâmetros extra.  
- Log detalhado por sessão.  
- Gestão de cancelamento.  
- Modo "eliminação segura" dos originais.  
- Persistência de estado entre sessões.

---

## Requisitos

- Windows com PowerShell 5.1 ou superior.  
- 7-Zip instalado.  

---

## Instalação

1. Instale o 7-Zip.  
2. Execute o script.  
3. Ajuste ExecutionPolicy se necessário.