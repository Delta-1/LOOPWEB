# LOOPWEB

Aplicativo de orçamento com catálogo, fotos de produtos, Google Sheets e Google Drive.

## Correção da sincronização de fotos

O front-end agora:

- comprime fotos de celular antes do envio;
- cria um nome único para cada nova imagem;
- permite trocar a foto pela tela de edição do produto;
- impede códigos duplicados no cadastro;
- aguarda a confirmação real da planilha antes de informar que salvou;
- evita que o navegador continue mostrando uma foto antiga em cache.

O arquivo `apps-script/Code.gs` contém o back-end compatível com esse fluxo. Ele usa
a aba `Produtos` (ou a primeira aba da planilha) e reconhece cabeçalhos em português
ou inglês para código, nome, preço, estoque, categoria, unidade e imagem.

## Publicação do Apps Script

1. Abra a planilha e acesse **Extensões → Apps Script**.
2. Substitua o conteúdo do arquivo `Code.gs` pelo conteúdo de `apps-script/Code.gs`.
3. Em **Implantar → Nova implantação**, selecione **Aplicativo da Web**.
4. Execute como proprietário e permita acesso a qualquer pessoa com o link.
5. Copie a URL terminada em `/exec` para as configurações do LOOPWEB.
6. Informe também, no LOOPWEB, o ID da pasta do Google Drive que receberá as fotos.

Ao atualizar uma implantação já existente, crie uma nova versão para que o código
novo seja realmente usado pela URL pública.
