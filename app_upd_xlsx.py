import streamlit as st
import pandas as pd

# Configuração da página do Streamlit
st.set_page_config(page_title="Central de Utilitarios", layout="wide")

# -----------------------------------------------------------------------------
# FUNCOES DE SUPORTE
# -----------------------------------------------------------------------------
def formatar_valor_firebird(valor, forcar_string=False):
    """Formata os valores de acordo com os tipos aceitos no Firebird."""
    if pd.isna(valor):
        return "NULL"
    if forcar_string or isinstance(valor, str):
        texto = str(valor).replace("'", "''")
        return f"'{texto}'"
    elif isinstance(valor, bool):
        return "1" if valor else "0"
    elif isinstance(valor, (int, float)):
        return str(valor)
    return f"'{str(valor)}'"

def processar_linhas_sql(df, tabela, cols_set, cols_where, set_fixo, where_fixo, dicionario_tipos):
    """Executa a varredura das linhas do DataFrame."""
    scripts = []
    for _, linha in df.iterrows():
        # 1. Montagem do Bloco SET
        partes_set = []
        for col in cols_set:
            deve_string = dicionario_tipos.get(col, False)
            valor_fmt = formatar_valor_firebird(linha[col], forcar_string=deve_string)
            partes_set.append(f"{col} = {valor_fmt}")
        if set_fixo:
            partes_set.append(set_fixo.strip())
        clausula_set = ", ".join(partes_set)
        
        # 2. Montagem do Bloco WHERE
        partes_where = []
        for col in cols_where:
            deve_string = dicionario_tipos.get(col, False)
            valor_fmt = formatar_valor_firebird(linha[col], forcar_string=deve_string)
            partes_where.append(f"{col} = {valor_fmt}")
        if where_fixo:
            partes_where.append(f"({where_fixo})")
        condicao_final = " AND ".join(partes_where)
        
        # Consolidação da linha
        sql = f"UPDATE {tabela} SET {clausula_set} WHERE {condicao_final};"
        scripts.append(sql)
        
    return "\n".join(scripts)

# -----------------------------------------------------------------------------
# MENU DE NAVEGACAO
# -----------------------------------------------------------------------------
st.sidebar.title("Menu de Utilitarios")
opcao = st.sidebar.radio(
    "Selecione uma ferramenta:",
    options=["Excel para Script SQL (Firebird)", "Utilitario 2", "Utilitario 3"]
)

# -----------------------------------------------------------------------------
# FERRAMENTA 1: EXCEL PARA FIREBIRD SQL
# -----------------------------------------------------------------------------
if opcao == "Excel para Script SQL (Firebird)":
    st.title("Excel para Firebird SQL Update")
    st.write("Filtre, mapeie tipos/nomes, insira dados fixos e gere seu script de atualizacao.")

    arquivo_upload = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=["xlsx"])

    if arquivo_upload is not None:
        try:
            df_original = pd.read_excel(arquivo_upload)
            todas_colunas = df_original.columns.tolist()
            st.success("Planilha carregada com sucesso!")
        except Exception as e:
            st.error(f"Erro ao ler o arquivo Excel: {e}")
            st.stop()

        # ETAPA 1: FILTRAR COLUNAS DA PLANILHA
        st.divider()
        st.subheader("Etapa 1: Selecionar Colunas Uteis")
        colunas_selecionadas = st.multiselect(
            "Selecione apenas as colunas da planilha que serao usadas no UPDATE ou no WHERE:",
            options=todas_colunas,
            default=todas_colunas
        )

        if not colunas_selecionadas:
            st.warning("Por favor, selecione ao menos uma coluna para prosseguir.")
        else:
            # ETAPA 2: MAPEAMENTO DE CABECALHOS (DE / PARA) E DEFINICAO DE TIPO
            st.divider()
            st.subheader("Etapa 2: Mapeamento de Colunas e Tipos (DE / PARA)")
            st.info("Informe o nome real da coluna no Firebird e marque se deseja forçar o campo a agir como Texto.")
            
            mapeamento_colunas = {}
            colunas_forcar_string = {}
            
            for col_original in colunas_selecionadas:
                c_nome, c_tipo = st.columns(2)
                with c_nome:
                    nome_banco = st.text_input(
                        f"Coluna: {col_original} -> PARA Banco:", 
                        value=col_original,
                        key=f"n_{col_original}"
                    ).strip().upper()
                    mapeamento_colunas[col_original] = nome_banco
                    
                with c_tipo:
                    st.write(" ")
                    st.write(" ")
                    forcar = st.checkbox("Forçar como String (Texto)", key=f"s_{col_original}")
                    colunas_forcar_string[nome_banco] = forcar

            # Aplicação dos novos nomes e exibição da prévia
            df_filtrado = df_original[colunas_selecionadas].rename(columns=mapeamento_colunas)
            colunas_finais_disponiveis = df_filtrado.columns.tolist()

            st.write("---")
            st.caption("Previa simplificada dos dados:")
            st.dataframe(df_filtrado.head(3), use_container_width=True)

            # ETAPA 3: CONFIGURACOES DO SCRIPT SQL
            st.divider()
            st.subheader("Etapa 3: Configuracoes do Script SQL")
            
            col_tab, col_set, col_where = st.columns([1, 1.2, 1.5])

            with col_tab:
                st.markdown("**Nome da Tabela**")
                nome_tabela = st.text_input("Tabela destino:", value="MINHA_TABELA").upper()

            with col_set:
                st.markdown("**Campos para Atualizar (SET)**")
                colunas_atualizar = st.multiselect("Colunas dinamicas (da Planilha):", options=colunas_finais_disponiveis)
                add_set_fixo = st.checkbox("Adicionar campos fixos no SET?")
                campos_set_fixos = st.text_input("Campos fixos (SQL puro):", placeholder="SITUACAO = 'A'") if add_set_fixo else ""

            with col_where:
                st.markdown("**Filtros de Busca (WHERE)**")
                colunas_where_excel = st.multiselect("Colunas para o filtro WHERE dinamico:", options=[c for c in colunas_finais_disponiveis if c not in colunas_atualizar])
                adicionar_where_fixo = st.checkbox("Adicionar condicoes fixas no WHERE?")
                clausula_where_fixa = st.text_input("Condicoes fixas adicionais:", placeholder="EMPRESA_ID = 1") if adicionar_where_fixo else ""

            # PROCESSAMENTO DO ARQUIVO SQL
            if st.button("Gerar Script SQL para Firebird", type="primary"):
                if not colunas_atualizar and not campos_set_fixos:
                    st.warning("Por favor, selecione ou digite pelo menos um campo para atualizar no bloco SET.")
                elif not colunas_where_excel and not clausula_where_fixa:
                    st.warning("Atencao: Sem filtros (WHERE), voce atualizara TODOS os registros do banco.")
                else:
                    script_final = processar_linhas_sql(
                        df=df_filtrado,
                        tabela=nome_tabela,
                        cols_set=colunas_atualizar,
                        cols_where=colunas_where_excel,
                        set_fixo=campos_set_fixos,
                        where_fixo=clausula_where_fixa,
                        dicionario_tipos=colunas_forcar_string
                    )

                    st.divider()
                    st.subheader("Script Gerado")
                    st.code(script_final, language="sql")
                    
                    st.download_button(
                        label="Baixar arquivo .sql",
                        data=script_final,
                        file_name="update_firebird.sql",
                        mime="text/sql"
                    )

# -----------------------------------------------------------------------------
# OUTROS UTILITARIOS
# -----------------------------------------------------------------------------
elif opcao == "Utilitario 2":
    st.title("Proxima Ferramenta")
    st.write("Pronto para receber codigo.")

elif opcao == "Utilitario 3":
    st.title("Proxima Ferramenta")
    st.write("Pronto para receber codigo.")
