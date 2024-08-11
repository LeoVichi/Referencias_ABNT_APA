# Gerador de Referências em ABNT e APA

Este projeto Python processa e formata referências bibliográficas nos estilos ABNT e APA 7ª edição. Ele é capaz de lidar com referências manuais, bem como buscar dados usando DOI e ISBN. O código precisa de ajustes, pois ainda não está reconhecendo alguns links DOI e alguns ISBNs.

## Funcionalidades

- **Processamento de referências bibliográficas**: Suporte a referências manuais, DOI e ISBN.
- **Formatação**: Formatação de referências nos estilos ABNT e APA 7ª edição.
- **Geração de documentos**: Salva as referências formatadas em arquivos `.docx`.

## Como Usar

1. **Clone o repositório:**

   ```bash
   git clone https://github.com/LeoVichi/Refer-ncias_ABNT_APA.git
   cd reference-formatter
   ```

2. **Instale as dependências:**

   Certifique-se de que o Python 3.x está instalado em seu sistema. Em seguida, instale as dependências usando o `pip`:

   ```bash
   pip install -r requirements.txt
   ```

3. **Prepare o arquivo de entrada:**

   Crie um arquivo `referencias.txt` no mesmo diretório do script, contendo as referências que deseja processar.

4. **Execute o Script:**

   ```bash
   python reference_formatter.py
   ```

5. **Visualize os Resultados:**

   O script gerará dois arquivos `.docx`, um para referências ABNT (`referencias_abnt.docx`) e outro para referências APA 7ª edição (`referencias_apa7.docx`).

## Requisitos

- Python 3.x
- Bibliotecas listadas no `requirements.txt`

## Licença

Este projeto é licenciado sob os termos da licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

## Autor

- **L3nny_P34s4n7**
- **Email:** contact@leonardovichi.com
