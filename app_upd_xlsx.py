import streamlit as st
import pandas as pd

# Configuracao da pagina do Streamlit
st.set_page_config(page_title="Central de Utilitarios", layout="wide")

# -----------------------------------------------------------------------------
# FUNCOES DE SUPORTE
# -----------------------------------------------------------------------------
def formatar_valor_firebird(valor, forcar_string=False, is_data=False, is_decimal=False):
    """Formata os valores de acordo com os tipos aceitos no Firebird."""
    if pd.isna(valor):
        return "NULL"
    if is_data and hasattr(valor, "strftime"):
        return f"'{valor.strftime('%Y-%m-%d')}'"
    if forcar_string or isinstance(valor, str):
        texto = str(valor).replace("'", "''")
        return f"'{texto}'"
    elif isinstance(valor, bool):
        return "1" if valor else "0"
    elif isinstance(valor, (int, float)):
        if is_decimal:
            return f"{valor:.2f}"
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor)
    return f"'{str(valor)}'"

def processar_linhas_sql(df, tabela, cols_set, cols_where, set_fixo, where_fixo, dicionario_tipos, cols_decimal=None):
    """Executa a varredura das linhas do DataFrame para gerar UPDATEs."""
    cols_data_auto = {c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])}
    scripts = []
    for _, linha in df.iterrows():
        partes_set = []
        for col in cols_set:
            deve_string = dicionario_tipos.get(col, False)
            is_data = col in cols_data_auto
            is_decimal = col in (cols_decimal or [])
            valor_fmt = formatar_valor_firebird(linha[col], forcar_string=deve_string, is_data=is_data, is_decimal=is_decimal)
            partes_set.append(f"{col} = {valor_fmt}")
        
        if set_fixo and set_fixo.strip() != "":
            partes_set.append(set_fixo.strip())
            
        clausula_set = ", ".join(partes_set)
        
        partes_where = []
        for col in cols_where:
            deve_string = dicionario_tipos.get(col, False)
            is_data = col in cols_data_auto
            valor_fmt = formatar_valor_firebird(linha[col], forcar_string=deve_string, is_data=is_data)
            partes_where.append(f"{col} = {valor_fmt}")
            
        if where_fixo and where_fixo.strip() != "":
            partes_where.append(f"({where_fixo.strip()})")
            
        condicao_final = " AND ".join(partes_where)
        
        sql = f"UPDATE {tabela} SET {clausula_set} WHERE {condicao_final};"
        scripts.append(sql)
        
    return "\n".join(scripts)

def processar_linhas_insert_sql(df, tabela, cols_insert, campos_fixos, dicionario_tipos, cols_decimal=None):
    """Executa a varredura das linhas do DataFrame para gerar INSERTs."""
    scripts = []
    lista_cols_fixas = []
    lista_valores_fixos = []
    cols_data_auto = {c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])}
    
    if campos_fixos and campos_fixos.strip() != "":
        pares = campos_fixos.split(",")
        for par in pares:
            if "=" in par:
                c, v = par.split("=", 1)
                lista_cols_fixas.append(c.strip().upper())
                lista_valores_fixos.append(v.strip())

    for _, linha in df.iterrows():
        colunas_finais = []
        valores_finais = []
        
        for col in cols_insert:
            deve_string = dicionario_tipos.get(col, False)
            is_data = col in cols_data_auto
            is_decimal = col in (cols_decimal or [])
            valor_fmt = formatar_valor_firebird(linha[col], forcar_string=deve_string, is_data=is_data, is_decimal=is_decimal)
            colunas_finais.append(col)
            valores_finais.append(valor_fmt)
            
        for col_fixa, val_fixo in zip(lista_cols_fixas, lista_valores_fixos):
            colunas_finais.append(col_fixa)
            valores_finais.append(val_fixo)
            
        str_colunas = ", ".join(colunas_finais)
        str_valores = ", ".join(valores_finais)
        
        sql = f"INSERT INTO {tabela} ({str_colunas}) VALUES ({str_valores});"
        scripts.append(sql)
        
    return "\n".join(scripts)

# -----------------------------------------------------------------------------
# MENU DE NAVEGACAO (Barra Lateral)
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ Utilitários")
opcao = st.sidebar.radio(
    "Selecione uma ferramenta:",
    options=[
        "📊 Excel para Script SQL (UPDATE)", 
        "➕ Excel para Script SQL (INSERT)", 
        "⚙️ Utilitário 3 (Vazio)"
    ]
)

# -----------------------------------------------------------------------------
# FERRAMENTA 1: EXCEL PARA FIREBIRD SQL (UPDATE)
# -----------------------------------------------------------------------------
if opcao == "📊 Excel para Script SQL (UPDATE)":
    st.title("📊 Excel para Firebird SQL Update")
    st.write("Filtre colunas da planilha, mapeie os campos para o banco, insira dados fixos e gere o script.")

    arquivo_upload = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=["xlsx"], key="up_update")

    if arquivo_upload is not None:
        df_original = pd.read_excel(arquivo_upload)
        todas_colunas = df_original.columns.tolist()
        st.success("Planilha carregada com sucesso!")

        st.divider()
        st.subheader("🎯 Etapa 1: Selecionar Colunas Úteis")
        colunas_selecionadas = st.multiselect(
            "Selecione apenas as colunas da planilha que serão usadas no UPDATE ou no WHERE:",
            options=todas_colunas,
            default=todas_colunas,
            key="ms_update"
        )

        if not colunas_selecionadas:
            st.warning("Por favor, selecione ao menos uma coluna para prosseguir.")
        else:
            st.divider()
            st.subheader("🔄 Etapa 2: Mapeamento de Colunas e Tipos (DE / PARA)")
            st.info("Informe o nome real da coluna no Firebird e marque se deseja forçar o campo a agir como Texto.")
            
            mapeamento_colunas = {}
            colunas_forcar_string = {}
            cols_decimal_update = st.multiselect("Colunas DECIMAL (2 casas):", options=colunas_selecionadas, key="decimal_update")
            
            for col_original in colunas_selecionadas:
                c_nome, c_tipo = st.columns(2)
                with c_nome:
                    nome_banco = st.text_input(
                        f"Coluna: {col_original} ➡️ PARA Banco:", 
                        value=col_original,
                        key=f"n_up_{col_original}"
                    ).strip().upper()
                    mapeamento_colunas[col_original] = nome_banco
                    
                with c_tipo:
                    st.write(" ")
                    st.write(" ")
                    forcar = st.checkbox("Forçar como String (Texto)", key=f"s_up_{col_original}")
                    colunas_forcar_string[nome_banco] = forcar

            df_filtrado = df_original[colunas_selecionadas].rename(columns=mapeamento_colunas)
            colunas_finais_disponiveis = df_filtrado.columns.tolist()

            st.write("---")
            st.caption("📋 Prévia simplificada dos dados com cabeçalhos atualizados:")
            st.dataframe(df_filtrado.head(3), use_container_width=True)

            st.divider()
            st.subheader("⚙️ Etapa 3: Configurações do Script SQL")
            
            col_tab, col_set, col_where = st.columns([1, 1.2, 1.5])

            with col_tab:
                st.markdown("**📋 Nome da Tabela**")
                nome_tabela = st.text_input("Tabela destino:", value="MINHA_TABELA", key="tab_update").upper()

            with col_set:
                st.markdown("**✏️ Campos para Atualizar (SET)**")
                colunas_atualizar = st.multiselect("Colunas dinâmicas (da Planilha):", options=colunas_finais_disponiveis, key="ms_set_din")
                campos_set_fixos = st.text_input(
                    "Campos fixos adicionais (Opcional):", 
                    placeholder="SITUACAO = 'A', DATA = 'NOW'",
                    key="txt_set_fixo"
                )

            with col_where:
                st.markdown("**🔍 Filtros de Busca (WHERE)**")
                colunas_where_excel = st.multiselect("Colunas para o filtro WHERE dinâmico:", options=[c for c in colunas_finais_disponiveis if c not in colunas_atualizar], key="ms_where_din")
                clausula_where_fixa = st.text_input(
                    "Condições fixas adicionais (Opcional):", 
                    placeholder="EMPRESA_ID = 1",
                    key="txt_where_fixo"
                )

            st.write("---")
            if st.button("🚀 Gerar Script SQL para Firebird", type="primary", key="btn_update"):
                if not colunas_atualizar and not campos_set_fixos.strip():
                    st.warning("Por favor, selecione ou digite pelo menos um campo para atualizar no bloco SET.")
                elif not colunas_where_excel and not clausula_where_fixa.strip():
                    st.warning("Atenção: Sem filtros (WHERE), você atualizará TODOS os registros do banco.")
                else:
                    script_final = processar_linhas_sql(
                        df=df_filtrado,
                        tabela=nome_tabela,
                        cols_set=colunas_atualizar,
                        cols_where=colunas_where_excel,
                        set_fixo=campos_set_fixos,
                        where_fixo=clausula_where_fixa,
                        dicionario_tipos=colunas_forcar_string,
                        cols_decimal=cols_decimal_update
                    )

                    st.divider()
                    st.subheader("📝 Script Gerado")
                    st.code(script_final, language="sql")
                    
                    st.download_button(
                        label="💾 Baixar arquivo .sql",
                        data=script_final,
                        file_name="update_firebird.sql",
                        mime="text/sql",
                        key="dl_update"
                    )

# -----------------------------------------------------------------------------
# FERRAMENTA 2: EXCEL PARA FIREBIRD SQL (INSERT)
# -----------------------------------------------------------------------------
elif opcao == "➕ Excel para Script SQL (INSERT)":
    st.title("➕ Excel para Firebird SQL Insert")
    st.write("Filtre colunas da planilha, faça o mapeamento e gere scripts de inserção (INSERT INTO).")

    arquivo_upload = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=["xlsx"], key="up_insert")

    if arquivo_upload is not None:
        df_original = pd.read_excel(arquivo_upload)
        todas_colunas = df_original.columns.tolist()
        st.success("Planilha carregada com sucesso!")

        st.divider()
        st.subheader("🎯 Etapa 1: Selecionar Colunas Úteis")
        cols_insert = st.multiselect(
            "Selecione as colunas da planilha que serão usadas no INSERT:",
            options=todas_colunas,
            default=todas_colunas,
            key="ms_insert"
        )

        if not cols_insert:
            st.warning("Por favor, selecione ao menos uma coluna para prosseguir.")
        else:
            st.divider()
            st.subheader("🔄 Etapa 2: Mapeamento de Colunas e Tipos (DE / PARA)")
            st.info("Informe o nome real da coluna no Firebird e marque se deseja forçar o campo a agir como Texto.")
            
            mapeamento_colunas = {}
            colunas_forcar_string = {}
            
            for col_original in cols_insert:
                c_nome, c_tipo = st.columns(2)
                with c_nome:
                    nome_banco = st.text_input(
                        f"Coluna: {col_original} ➡️ PARA Banco:", 
                        value=col_original,
                        key=f"n_ins_{col_original}"
                    ).strip().upper()
                    mapeamento_colunas[col_original] = nome_banco
                    
                with c_tipo:
                    st.write(" ")
                    st.write(" ")
                    forcar = st.checkbox("Forçar como String (Texto)", key=f"s_ins_{col_original}")
                    colunas_forcar_string[nome_banco] = forcar

            df_filtrado = df_original[cols_inserir].renomear(colunas=mapeamento_colunas)
            colunas_finais_disponiveis = df_filtrado.colunas.listar()

            São.escrever("---")
            São.legenda("📋 Prévia simplificada dos dados com cabeçalhos atualizados:")
            São.quadro de dados(df_filtrado.cabeça(3), usar_largura_do_contêiner=Verdadeiro)

            São.divisor()
            São.subtítulo("⚙️ Etapa 3: Configurações do Script SQL")

            col_tab, col_fixos = st.colunas([1, 1,5])

            com col_tab:
                São.redução("**📋 Nome da Mesa**")
                nome_tabela = st.entrada_texto("Tabela destino:", valor="MINHA_TABELA", chave="tab_insert").superior()

            com col_fixos:
                São.redução("**⚡ Campos Fixos Adicionais (Opcional)**")
             ## st.info("Formato: CAMPO1 = VALOR1, CAMPO2 = VALOR2")
                campos_fixos = st.entrada_texto(
                    "Campos fixos:", 
                    espaço reservado="ID_STATUS = 'A', DATA_CRIACAO = AGORA",
                    chave="txt_fixos_insert"
                )

            São.escrever("---")
            São.subtítulo("💰 Colunas monetárias")
            cols_decimal_insert = st.multisseleção("Colunas DECIMAL (2 casas):", opções=colunas_finais_disponiveis, chave="inserção_decimal")

            São.escrever("---")
            se São.botão("🚀 Gerar Script SQL para Firebird", tipo="primário", chave="btn_inserir"):
                se não campos_fixos.tira() e cols_insert:
                    passar  # Campos fixos são opcionais
                
                script_final = processar_linhas_insert_sql(
                    df=df_filtrado,
                    tabela=nome_tabela,
                    cols_insert=colunas_finais_disponiveis,
                    campos_fixos=campos_fixos,
                    dicas_tipos=colunas_forcar_string,
                    cols_decimal=cols_decimal_inserir
                )

                São.divisor()
                São.subtítulo("📝 Roteiro Gerado")
                São.código(script_final, idioma="sql")
                
                São.botão_download(
                    rótulo="💾 Baixar arquivo .sql",
                    dados=script_final,
                    nome_arquivo="inserir_firebird.sql",
                    mimica="texto/sql",
                    chave="dl_inserir"
                )
