# 🦁 Organizador_Informacoes_Declaracao_IR (Excel)

Origem: Bootcamp SOA / DIO - Excel+AI.

## 📌 Sobre o Projeto
Este projeto tem como objetivo organizar, de forma visual e estruturada, as principais informações necessárias para a declaração de imposto de renda de pessoa física.

A proposta foi desenvolver uma ferramenta no Excel que permita centralizar dados cadastrais, bancários e rendimentos em um único arquivo, com navegação intuitiva, interativas entre si e com validações que auxiliam no correto preenchimento das informações.


## 🧩 Estrutura da Solução

O arquivo é composto por três planilhas interativas, que se comunicam entre si por meio de botões e hiperlinks, facilitando a navegação:

- Dados do Titular
Tela destinada as informações pessoais, com aplicação de máscaras e validações para garantir padronização e consistência dos dados.

- Dados e informes Bancários
Área para registro de até três contas bancárias, com consolidação automática do saldo total declarado.

- Notas bancárias e Extratos de holerites\
Planilha para organização das entradas financeiras, categorizadas por origem (holerite, CNPJ e freelance), com indicação de mês e valor recebido.

O projeto não tem foco em cálculos complexos, mas sim em organização, padronização e gerenciamento estruturado das informações.

## 🎨 Técnicas Visuais e Recursos Utilizados

Neste projeto, o foco principal foi a organização visual e a usabilidade, explorando recursos do Excel para tornar a experiência mais clara e intuitiva:

- Criação de menu lateral de navegação
- Uso de botões com hiperlinks internos entre planilhas
- Links externos para acesso rápido a sites relevantes
- Padronização visual e identidade do layout (cores, fontes, alinhamentos e espaçamentos)
- Formatação com uso de células destacadas para entrada de dados
- Aplicação de máscaras predefinidas, como: CPF no formato 000.000.000-00 / Telefone no formato (00) 0000-0000
  
Validação de dados:
- Incluindo listas suspensas com dados pre determinado.
- restringindo e demilitando a entrada de dados em celulas para apenas de números e limite de caracteres (ex: CPF com 11 dígitos) usando a formula:
```
=E(ÉNÚM(D7);NÚM.CARACT(D7)=11)
```

- Implementação de uma função auxiliar em VBA para ajuste de tamanho e alinhamento de ícones, suprindo uma limitação nativa do Excel
```
Sub MoverIconeParaPosicao()
    Dim shp As Shape
    Dim ws As Worksheet
    Dim nomeIconeProcurado As String
    Dim novaPosicaoX As Double
    Dim novaPosicaoY As Double
    
    ' Defina a planilha atual
    Set ws = ActiveSheet
    
    ' Defina o nome do Ã­cone que vocÃª quer mover (exato, como aparece no Excel)
    nomeIconeProcurado = "Ãcone 1" ' <-- Troque aqui pelo nome do seu Ã­cone
    
    ' Defina a posiÃ§Ã£o desejada
    novaPosicaoX = 100 ' PosiÃ§Ã£o X em pontos
    novaPosicaoY = 50  ' PosiÃ§Ã£o Y em pontos
    
    ' Procura pelo Ã­cone na planilha
    For Each shp In ws.Shapes
        If shp.Name = nomeIconeProcurado Then
            ' Move o Ã­cone para a nova posiÃ§Ã£o
            shp.Left = novaPosicaoX
            shp.Top = novaPosicaoY
            MsgBox "Ãcone '" & nomeIconeProcurado & "' movido com sucesso!", vbInformation
            Exit Sub
        End If
    Next shp
    
    ' Se nÃ£o encontrar
    MsgBox "Ãcone '" & nomeIconeProcurado & "' nÃ£o encontrado.", vbExclamation
End Sub
```

## 🔢 Funções Utilizadas

- SOMA() para cálculo de consolidação simples de valores bancários.

## 🗂 Uso de Planilha de Apoio

Foi criada uma planilha auxiliar contendo:
- Uma lista de bancos para uso na validação de dados

Essa estrutura permite:
- Separação lógica
- Manutenção facilitada
- Escalabilidade do modelo

## 🎯 Objetivo do Projeto

O objetivo principal foi praticar conceitos de organização, validação e apresentação visual de dados no Excel, criando uma ferramenta funcional para uso real, com foco em formatacao, estilos, clareza nas informacoes, navegação simples e interabilidade, padronização.

## 👩‍💻 Autora

Andrielle Cunha - Intusiasta de Dados

