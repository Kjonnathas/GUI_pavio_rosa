# GUI_pavio_rosa
 
 Etapas do projeto:
 
 1. Criação da modelagem entidade-relacionamento do banco de dados utilizando o StarUML;
 
 2. Criação do banco de dados e das tabelas no SQL Server;
 
 3. Desenvolvimento do projeto em Python, tanto o front-end quanto o back-end:
 
  3.1. Para o front-end foi utilizada a lib do Customtkinter e o Tkinter;
  
  3.2. Para o back-end foram utilizadas algumas libs para auxiliar no desenvolvimento, tais como:
  
   3.2.1. lib win32com.client para disparar e-mail quando o usuário esquecer a senha;
   
   3.2.2 lib pandas para ler, tratar e exportar as informações puxadas do banco de dados;
   
   3.2.3. lib numpy para conversão do tipo de dado;
   
   3.2.4. lib webbrowser para abrir o navegador e redirecionar o usuário para sua página de e-mail;
   
   3.2.5. lib pyodbc para realizar a integração e conexão ao banco de dados;
   
   3.2.6. lib datetime para conversão e tratamento de datas;
   
   3.2.7 lib time para fazer o programa aguardar alguns segundos antes de executar os próximos comandos;
   
   3.2.8. lib warnings para ignorar os avisos e não poluir o terminal;
   
   3.2.9. lib os para auxiliar na interação com o sistema e pastas do computador;
   
   3.2.10. lib re para auxiliar na identificação de padrões fornecidos pelo usuário.
   
4. Criação de um ambiente virtual para instalar apenas as bibliotecas utilizadas no desenvolvimento do projeto, de modo a não deixar o programa extremamente pesado;

5. Transformação do arquivo .py em .exe para que possa rodar em qualquer máquina. Neste caso utilizei tanto alguns comandos pelo prompt quanto o auto-py-to-exe para realizar a transformação do arquivo.
