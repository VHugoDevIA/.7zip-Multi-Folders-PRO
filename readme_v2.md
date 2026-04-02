# ZIP Múltiplas Pastas - PRO v2

Aplicação **PowerShell** com interface gráfica Windows Forms para automatizar a compressão de múltiplas pastas usando o motor do 7-Zip (7z.exe).  
Versão básica focada na compressão simples de pastas em formato ZIP.

---

## Funcionalidades principais

- Interface gráfica (WinForms) para seleção e gestão de várias pastas numa só operação.  
- Suporte para formato `zip` com nível de compressão configurável (1-9).  
- Geração automática de nomes de arquivo únicos se já existir (sufixo incremental).  
- Possibilidade de arrastar e largar pastas para a lista.  
- Botões para adicionar/remover pastas da lista.  
- Campo para escolher a pasta de destino.  
- Execução da compressão com barra de progresso.  

---

## Requisitos

- Windows com suporte a PowerShell.  
- 7-Zip instalado em `C:\Program Files\7-Zip\7z.exe`.  
- Permissões para executar scripts PowerShell.

---

## Instalação

1. Certifique-se de que o 7-Zip está instalado no caminho padrão.  
2. Execute o script `zip_multiplas_pastas_pro_v2.ps1` com PowerShell.  
3. Ajuste a política de execução se necessário: `Set-ExecutionPolicy RemoteSigned`.

---

## Utilização

1. Abra o script no PowerShell.  
2. Adicione pastas à lista usando o botão "Adicionar pasta" ou arrastando.  
3. Escolha a pasta de destino.  
4. Clique em "Comprimir" para iniciar.  

---

## Notas

Esta é uma versão básica sem funcionalidades avançadas como encriptação ou logs.