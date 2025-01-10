# MailTasker VBA  Gerenciador de Tarefas em VBA
Automatize tarefas no Excel com notificações por e-mail via Outlook.

## 📋 Funcionalidades

- Gerenciamento de tarefas com status (Concluído, Pendente).
- Envio automático de e-mails de notificação para prazos vencidos.
- Interface configurável diretamente no Excel.

## ⚙️ Pré-requisitos

- **Microsoft Excel** com suporte a macros habilitado.
- **Microsoft Outlook** configurado no sistema.
- Permissões para executar código VBA.

## 🛠️ Configuração

1. Baixe ou clone o repositório:
   ```bash
   git clone https://github.com/Shakalinux/MailTasker-vba.git
   ```

2. Abra o arquivo Excel (.xlsm) no Microsoft Excel.

3. Certifique-se de que as macros estejam habilitadas:
   - Vá em **Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Configurações de Macro**.
   - Selecione "Habilitar todas as macros".

4. Personalize a tabela inicial com suas tarefas e e-mails.

## 🚀 Uso

1. Preencha a tabela com as informações:
   - **Tarefa**: Descrição da tarefa.
   - **Responsável**: E-mail do responsável pela tarefa.
   - **Data de Conclusão**: Data esperada para finalizar a tarefa.
   - **Status**: Será atualizado automaticamente.

2. Execute o VBA para:
   - Verificar tarefas pendentes com base na data.
   - Enviar notificações automáticas por e-mail via Outlook.

## 📧 Envio de E-mails

O envio de e-mails utiliza o Microsoft Outlook para disparar mensagens diretamente para os endereços registrados na tabela. Certifique-se de que o Outlook esteja configurado no computador.


## 📝 Licença

Este projeto está licenciado sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## Demostração da interface e código
<div style="width:100%; overflow:hidden; max-width:600px;">
  <div style="display: flex; transition: transform 0.5s ease;">
    <img src="https://i.postimg.cc/d1Sb7VJk/tabela.png" alt="Imagem 1" style="width:100%; flex-shrink: 0; border: 5px solid black;">
    <img src="https://i.postimg.cc/y6DdLJgG/codigo1.png" alt="Imagem 2" style="width:100%; flex-shrink: 0; border: 5px solid black;">
    <img src="https://i.postimg.cc/h4NG6SVt/codigo2.png" alt="Imagem 3" style="width:100%; flex-shrink: 0; border: 5px solid black;">
     <img src="https://i.postimg.cc/pXvwp4HB/codigo3.png" alt="Imagem 3" style="width:100%; flex-shrink: 0; border: 5px solid black;">
  </div>
</div>





















