# ZIP Múltiplas Pastas - PRO v3

Aplicação **PowerShell** com interface gráfica Windows Forms para automatizar a compressão de múltiplas pastas usando o motor do 7-Zip (7z.exe).  
Introduz suporte para formatos ZIP e 7Z, e log detalhado na interface.

---

## Funcionalidades principais

- Interface gráfica (WinForms) para seleção e gestão de várias pastas numa só operação.  
- Suporte para formatos `zip` e `7z`, com nível de compressão configurável.  
- Geração automática de nomes de arquivo únicos se já existir (sufixo incremental).  
- Possibilidade de arrastar e largar pastas para a lista.  
- Botões para adicionar/remover pastas da lista.  
- Campo para escolher a pasta de destino e formato do arquivo.  
- Log detalhado na interface durante a compressão.  
- Execução da compressão com barra de progresso.  

---

## Requisitos

- Windows com suporte a PowerShell.  
- 7-Zip instalado em `C:\Program Files\7-Zip\7z.exe`.  
- Permissões para executar scripts PowerShell.

---

## Instalação

1. Certifique-se de que o 7-Zip está instalado no caminho padrão.  
2. Execute o script `zip_multiplas_pastas_pro_v3.ps1` com PowerShell.  
3. Ajuste a política de execução se necessário: `Set-ExecutionPolicy RemoteSigned`.

---

## Utilização

1. Abra o script no PowerShell.  
2. Selecione o formato (ZIP ou 7Z).  
3. Adicione pastas à lista usando o botão "Adicionar pasta" ou arrastando.  
4. Escolha a pasta de destino.  
5. Clique em "Comprimir" para iniciar e ver o log.  

---

## Notas

Adiciona suporte a 7Z e log na interface para melhor feedback.