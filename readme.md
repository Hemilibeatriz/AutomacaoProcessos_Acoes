Automação de Processo:
    Código para acompanhar o valor de algumas ações na B3. O código entra no Google Drive, puxa os valores das ações (pré estabelecidas para ele),com o uso do Selenium, alimenta uma planilha com a data e o valor de cada ação. E envia por email, as cotações através do OutLook.

Observação:
Você PRECISA do driver do Chrome (ChromeDriver) para que o Selenium possa funcionar bem. Para instalá-lo basta saber qual a versão do seu Google Chrome.

Para verificar qual a versão do navegador Google Chrome, clique no menu na barra de ferramentas Ajuda e selecione “Sobre o Google Chrome”. O número da versão atual aparecerá.

Pesquise na internet por "chromedriver Version 97.0.4692.71" (coloque a versão que voce encontrou na etapa anterior).

O primeiro link deverá ser este: https://chromedriver.chromium.org/downloads clique nele e escolha a versão correta.

Com o ChromeDriver baixado, escolha um local para estalação. Este local deve ver um PATH, para saber todos os seus PATHS vá no CMD e digite echo %path%

No meu caso, ele será instalado em C:\Users\hemil\AppData\Local\Programs\Python\Python39.

Sempre que der erro de versão no seu ChromeDriver, você deve atualiza-lo :)