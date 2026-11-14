# SISACP Web — Sistema de Acompanhamento e Controle de Peso

Versão web (rodando no seu próprio computador) do seu controle em Excel.
Nada vai para a internet: os dados ficam num único arquivo, `controle_peso.db`.

---

## 1. A ideia em uma figura

No Excel, **dados, cálculos, gráficos e botões vivem todos na mesma planilha**.
No sistema web, cada coisa tem o seu lugar:

```
  Você (navegador)
        │  clica, digita
        ▼
  ┌───────────────┐   app.py          -> telas, formulários, gráficos      (era: PRINCIPAL + Coletas)
  │  INTERFACE    │
  └───────┬───────┘
          │ pede números
  ┌───────▼───────┐   calculos.py     -> as "fórmulas" (IMC, %, metas)     (era: células =M7/M5...)
  │   REGRAS      │
  └───────┬───────┘
          │ lê / grava
  ┌───────▼───────┐   database.py     -> único arquivo que fala com o banco
  │  ACESSO AO BD │
  └───────┬───────┘
          │
  ┌───────▼───────┐   schema.sql      -> definição das tabelas             (era: as abas de dados)
  │ BANCO SQLite  │   controle_peso.db   (os dados em si)
  └───────────────┘

  importar_excel.py  -> roda UMA vez, copia o histórico do .xlsm para o banco
```

Vantagem prática: se um gráfico estiver errado, você mexe em `app.py`.
Se uma conta estiver errada, em `calculos.py`. Os dados nunca são tocados.

---

## 2. Instalação (Windows), passo a passo

**Passo 0 — Python.** Instale o Python 3.11 ou superior em https://www.python.org/downloads/
(na primeira tela do instalador, marque **"Add python.exe to PATH"**).

**Passo 1 — Abra um terminal na pasta do projeto.**
No Explorador de Arquivos, entre na pasta `sisacp`, clique na barra de endereço, digite `cmd` e Enter.

**Passo 2 — Crie o "ambiente virtual"** (uma caixa isolada para as bibliotecas do projeto):
```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```
Quando o terminal mostrar `(.venv)` no começo da linha, o ambiente está ativo.

**Passo 3 — Copie seu Excel** (`01 - Controle de Peso_v3.xlsm`) para dentro da pasta `sisacp`.

**Passo 4 — Importe o histórico** (uma única vez):
```
python importar_excel.py "01 - Controle de Peso_v3.xlsm"
```
Ele mostra quantas linhas gravou em cada tabela. Isso cria o arquivo `controle_peso.db`.

**Passo 5 — Abra o sistema:**
```
streamlit run app.py
```
O navegador abre em `http://localhost:8501`. Para fechar, `Ctrl + C` no terminal.

> Nos dias seguintes basta repetir só `.venv\Scripts\activate` e `streamlit run app.py`
> (ou dar duplo clique em `iniciar.bat`, que faz isso por você).

---

## 3. Uso no dia a dia

| Quero...                         | Vá em...               |
|----------------------------------|------------------------|
| Ver o painel (gráficos, metas)   | **📊 Painel**          |
| Lançar peso, água, treino, comida| **➕ Registrar**       |
| Trocar dieta / metas / itens     | **🍽️ Plano alimentar** |
| Ver remédios e custo             | **💊 Medicamentos**    |
| Mudar altura, meta de peso/água  | **👤 Perfil**          |

No Painel, o seletor **"Dia analisado"** substitui o filtro "Do dia?" das tabelas dinâmicas.
**Não existe mais "Atualizar tudo"**: toda tela lê o banco na hora.

---

## 4. Mapa: de onde veio cada coisa do Excel

| No Excel                                   | No sistema                                        |
|--------------------------------------------|---------------------------------------------------|
| Aba *Controle de Peso* (Tabela8)           | tabela `pesagem` + view `vw_pesagem` (IMC)        |
| Aba *Controle de Hidratação* (Tabela11)    | tabela `hidratacao`                               |
| Aba *Controle Atividades Fisicas* (Tabela7)| tabela `atividade_fisica`                         |
| Aba *Controle Alimentação*                 | tabela `consumo` + view `vw_consumo`              |
| *Apoio* > Tabela5 (itens com macros)       | tabela `item_plano`                               |
| *Apoio* > Tabela6 (metas por refeição)     | tabela `meta_refeicao`                            |
| Aba *Medicamentos*                         | tabela `medicamento`                              |
| Aba *Cardápio Semanal*                     | tabela `cardapio_item`                            |
| Dados de referência (nome, altura, meta)   | tabela `perfil`                                   |
| Macros VBA `Cadastrar_Peso` etc.           | botões da página **Registrar** (fazem `INSERT`)   |
| Tabelas dinâmicas + "Atualizar tudo"       | consultas SQL (`SUM`, `GROUP BY`) feitas ao abrir |
| PROCV dos nutrientes                       | view `vw_consumo` (`JOIN` entre consumo e item)   |

---

## 5. Conceitos novos, em linguagem de Excel

* **Banco de dados** = um arquivo com várias "tabelas" (como abas), mas cada tabela tem colunas
  de tipo fixo e regras (ex.: peso não pode ser negativo).
* **Tabela** = o equivalente a uma Tabela do Excel (Ctrl+T). Cada linha = um registro.
* **Chave primária (`id`)** = número único de cada linha. Serve para apontar para ela
  (é o "Cod Item" do seu Excel).
* **Chave estrangeira** = coluna que aponta para outra tabela. `consumo.item_plano_id` aponta
  para `item_plano.id`: o consumo guarda só *qual item* você comeu; calorias e macros são buscadas lá.
* **SQL** = a linguagem para pedir dados. `SELECT SUM(volume_ml) FROM hidratacao WHERE data='2026-11-05'`
  é o `SOMASE` do banco.
* **View** = consulta salva com nome; funciona como uma coluna calculada (ex.: IMC).
* **Transação** = "ou grava tudo ou não grava nada". Usada na importação.
* **Por que não gravar IMC, totais e %?** Porque se você corrigir uma altura ou meta, tudo que é
  *calculado* se atualiza sozinho; o que foi *gravado* ficaria errado para sempre.

---

## 6. Cuidados

**Backup:** o sistema inteiro de dados é o arquivo `controle_peso.db`. Copie-o de vez em quando
(para um pendrive ou nuvem). Para restaurar, coloque a cópia de volta na pasta.

**Nunca rode `importar_excel.py --recriar` depois de começar a usar o sistema:** ele apaga o banco
e refaz só com o que está no Excel. Sem `--recriar`, o script recusa-se a rodar se o banco já existe.

**Abrir no celular (mesmo Wi-Fi):**
```
streamlit run app.py --server.address 0.0.0.0
```
No celular, acesse `http://IP-DO-PC:8501` (descubra o IP com `ipconfig`). O Windows pode pedir
permissão no firewall. Só funciona dentro da sua rede doméstica; não exponha à internet.

---

## 7. O que NÃO foi migrado (de propósito)

* Abas ocultas **Lista Alimentos** e **Cont Alimentação**: versão antiga do controle (fev/2026).
* **Gráficos e tabelas dinâmicas**: foram refeitos do zero a partir dos dados.
* Colunas **SEMANA 1–4** da aba Medicamentos (marcações semanais).
* Aba *Imagens* e os ícones dos botões.

## 8. Diferenças em relação ao Excel (e por quê)

* **Rosca de redução:** o Excel exibia `1,1 / 24,2kg`, mas o "24,2" estava **digitado dentro da
  fórmula** (`="…/ 24,2kg"`). A redução necessária real é 20,4 kg (120,4 − 100). Aqui o valor é calculado.
* **Idade:** o Excel aproximava com 365,25 dias e 30,5 dias/mês; aqui é calendário exato.
* **"Obesidades Grau 2"** (com "s") virou **"Obesidade Grau 2"**.
* **Meta de calorias "até a última refeição":** soma das metas das refeições na ordem do dia
  (Café → Almoço → Lanche → Jantar → Ceia) até a última que você registrou.
* Cada consumo agora aceita **porções** (0,5 = metade), algo que o Excel não permitia.

## 9. Problemas comuns

| Sintoma                                   | Solução                                                        |
|-------------------------------------------|----------------------------------------------------------------|
| `'python' não é reconhecido`              | Reinstale o Python marcando "Add to PATH"; abra novo terminal. |
| `'streamlit' não é reconhecido`           | Ative o ambiente: `.venv\Scripts\activate`.                    |
| "Banco de dados não encontrado"           | Rode o Passo 4 (importação).                                   |
| "O banco já existe..." ao importar        | É a trava de segurança. Só use `--recriar` se quiser recomeçar.|
| Painel zerado                             | O dia selecionado não tem registros; troque o "Dia analisado". |

## 10. Próximos passos possíveis

Exportar para Excel/PDF; gráfico de gasto calórico (os campos `kcal_gastas` e `distancia_km` já
existem); lembretes de medicamento; múltiplos perfis (adicionar `perfil_id` nas tabelas);
migrar para PostgreSQL se um dia for usado por mais pessoas.
