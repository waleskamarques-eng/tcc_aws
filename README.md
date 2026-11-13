# Startup NuvemTech — Arquitetura Serverless na AWS

**Trabalho de Conclusão de Curso (TCC) | AWS re/Start e No Code — Escola da Nuvem | 2026**

## Sobre o projeto

A Startup NuvemTech é um case de gestão segura de documentos digitais, desenvolvido em equipe durante a formação da Escola da Nuvem.

O projeto apresenta uma arquitetura serverless na AWS para armazenamento, consulta e preservação de documentos, considerando segurança, escalabilidade, isolamento entre usuários e otimização de custos.

O cenário prevê o recebimento de aproximadamente 50 mil arquivos por mês, com retenção de longo prazo e preparação do acervo para futuras aplicações de inteligência artificial.

## Tecnologias e serviços AWS

- **Amazon S3:** armazenamento privado de documentos.
- **Amazon API Gateway:** disponibilização de endpoints HTTPS.
- **AWS Lambda:** processamento de requisições e regras de negócio.
- **Amazon Cognito:** autenticação e gerenciamento de usuários.
- **AWS IAM:** controle de permissões e princípio do menor privilégio.
- **AWS KMS:** gerenciamento de chaves de criptografia.
- **S3 Glacier Deep Archive:** armazenamento de longo prazo.
- **Amazon CloudWatch e AWS CloudTrail:** monitoramento e auditoria.

## Arquitetura da solução

A arquitetura utiliza autenticação com tokens JWT, processamento serverless e armazenamento privado no Amazon S3.

Os documentos são organizados por identificador de usuário, com validação de acesso no backend e utilização de URLs pré-assinadas para acesso temporário.

Uma política de ciclo de vida prevê a transição dos arquivos para armazenamento de menor custo após 365 dias, sem exclusão automática.

## Apresentação do projeto

O material completo contém 15 páginas, incluindo:

- Contexto e requisitos do case.
- Diagrama da arquitetura AWS.
- Serviços utilizados e decisões técnicas.
- Estratégias de segurança e preservação.
- Análise estimada de custos.
- Evidências práticas do portal NuvemTech.

**Arquivo:** [Visualizar apresentação completa — Startup NuvemTech (PDF)](APRESENTA%C3%87%C3%83O%20-%20A-Jornada-da-Startup-NuvemTech%20-%202.pdf)

## Equipe

Projeto desenvolvido em grupo por:

- Alexandre Ferreira Barbosa
- Ezequiel Elisel Farias Carrizo
- Leonardo Sena de Jesus
- Mateus de Oliveira Santos
- Sergio Adriano Faria
- Waleska Pereira Marques

## Formação

**Escola da Nuvem — AWS re/Start, Inteligência Artificial e No Code**

Formação de 300 horas, com atividades práticas e desenvolvimento de competências em computação em nuvem e tecnologias AWS.

---

*Projeto acadêmico desenvolvido para fins educacionais e de portfólio profissional.*
