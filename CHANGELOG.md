# Changelog

Todas as mudanças relevantes neste projeto serão documentadas neste arquivo.

O formato segue as diretrizes de [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/).

## [2.0.0] - 2025-05-03

### Novidades
- Refatoração completa do `reference_formatter.py` para melhorar legibilidade e modularidade.
- Implementação de logging com `logging.error()` e saída padronizada para `error_log.txt`.
- Suporte a fallback de metadados ISBN via **OpenLibrary API**, além do `isbnlib`.
- Suporte expandido a DOIs que usam **DataCite** (ex: Zenodo), além do Crossref.
- Títulos completamente em letras maiúsculas agora são automaticamente *normalizados* para *Title Case*.
- Configuração mais precisa de estilos `.docx`, com margens definidas em centímetros.
- Identificação de DOIs e ISBNs com expressões regulares robustas (`DOI_RE`, `ISBN_RE`).
- Override manual para DOIs específicos como `10.5281/zenodo.5829447`.

### Melhorias
- Formatação dos nomes dos autores ABNT e APA mais precisa, incluindo uso de iniciais no estilo APA 7.
- Código modularizado com funções mais claras e reutilizáveis.
- Melhoria no tratamento de títulos compostos com subtítulos (`:` em vez de `-`).
- Uso consistente de `str.title()` e `str.capitalize()` para normalizar nomes e títulos.
- Tratamento mais robusto de erros ao buscar metadados externos (Crossref, isbnlib, OpenLibrary).

### Remoções
- Funções duplicadas e redundantes foram eliminadas ou unificadas.
- Remoção de prints diretos para depuração no terminal; agora tudo é registrado via log.

### Alterações técnicas
- Dependências mantidas: `isbnlib`, `habanero`, `docx`, `requests`
- Novo padrão de codificação: Python 3.10+ com tipagem opcional (`-> str`, etc.)
- Compatível com ambientes virtuais (`venv`) e pronto para deploy com `requirements.txt`.

---

## [1.0.0] - 2025-05-03

- Primeira versão funcional com suporte a:
  - Entrada manual, DOI e ISBN
  - Formatação em estilos **ABNT** e **APA 7ª edição**
  - Geração automática de arquivos `.docx` com cabeçalho e formatação
  - Extração básica de metadados via `isbnlib` e `habanero`
