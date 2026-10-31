/**
 * Mapa de colunas da planilha "PROTOCOLO DE DOCS E MANIFESTAÇÕES".
 *
 * REGRA DE OURO: o app só escreve nas colunas listadas em `fields`.
 * As colunas em `formulas` nunca são tocadas.
 *
 * `editableUntil` é a última linha que ainda está dentro dos intervalos
 * liberados para edição (protectedRanges) das abas protegidas. Acima dessa
 * linha a planilha está totalmente bloqueada.
 */

export const SHEET_NAMES = {
  entrada: "Entrada de Documentos",
  // Atenção: o nome real da aba começa com um espaço.
  saida: " Saída e Produção",
  manifestacao: "Dados da Manifestação",
  proposta: "Proposta de Recompensa",
  cedmu: "Controle CEDMU",
  listas: "LISTAS",
};

/** Cabeçalho em `headerRow`; primeiro registro em `firstDataRow`. */
export const TABLES = {
  entrada: {
    key: "entrada",
    label: "Entrada",
    sheet: SHEET_NAMES.entrada,
    headerRow: 4,
    firstDataRow: 5,
    /** Coluna de data: referência para achar a última linha usada. */
    dateCol: "C",
    /** Nº do protocolo (preenchido pela planilha, só leitura). */
    protocolCol: "A",
    /** Nunca escrever nestas colunas. */
    formulas: ["A", "B", "D", "E", "F", "G", "N", "R", "S"],
    lastRowGuard: 30000,
    editableUntil: 15170,
    fields: [
      { key: "data", label: "Data do documento", col: "C", type: "date", required: true },
      {
        key: "tipo", label: "Tipo de documento", col: "H", type: "select",
        list: "tipoDocEntrada", required: true,
      },
      { key: "numDoc", label: "Nº do documento", col: "I", type: "text", placeholder: "Ex.: 142/2026 ou s/n" },
      { key: "docRef", label: "Documento de referência", col: "J", type: "text", list: "tipoDocEntrada" },
      { key: "numDocRef", label: "Nº do doc. de referência", col: "K", type: "text", placeholder: "REDS, processo, SICOR..." },
      { key: "origem", label: "Origem", col: "L", type: "text", list: "unidade", placeholder: "Unidade / órgão de origem" },
      { key: "remetente", label: "Remetente / manifestante", col: "M", type: "text" },
      { key: "envolvido", label: "Envolvido(s)", col: "O", type: "text", span: 2 },
      { key: "anexos", label: "Anexos", col: "P", type: "text", placeholder: "Quantidade / descrição" },
      { key: "arquivos", label: "Arquivos (link ou pasta)", col: "Q", type: "text" },
      { key: "comunicado", label: "Comunicado", col: "T", type: "text" },
      { key: "dataComunicado", label: "Data do comunicado", col: "U", type: "date" },
      { key: "destino", label: "Destino", col: "V", type: "text" },
      { key: "via", label: "Via", col: "W", type: "text" },
      { key: "dataRecebimento", label: "Data de recebimento", col: "X", type: "date" },
      { key: "obs", label: "Observações", col: "Y", type: "textarea", span: 2 },
    ],
  },

  saida: {
    key: "saida",
    label: "Saída",
    sheet: SHEET_NAMES.saida,
    headerRow: 3,
    firstDataRow: 4,
    dateCol: "D",
    protocolCol: "A",
    formulas: ["A", "C", "E", "F", "G", "H", "R", "S"],
    lastRowGuard: 30000,
    editableUntil: 16194,
    fields: [
      { key: "adj", label: "Adj.", col: "B", type: "text", placeholder: "Aditamento / anexo" },
      { key: "data", label: "Data do documento", col: "D", type: "date", required: true },
      {
        key: "tipo", label: "Tipo de documento", col: "I", type: "select",
        list: "tipoDocSaida", required: true,
      },
      { key: "assunto", label: "Assunto", col: "J", type: "text", span: 2 },
      { key: "envolvido", label: "Envolvido(s)", col: "K", type: "text", span: 2 },
      { key: "docRef", label: "Documento de referência", col: "L", type: "text" },
      { key: "numDocRef", label: "Nº do doc. de referência", col: "M", type: "text" },
      { key: "orgaoRef", label: "Órgão / unidade de referência", col: "N", type: "text", list: "unidade" },
      { key: "municipio", label: "Município", col: "O", type: "text" },
      { key: "anexos", label: "Anexos", col: "P", type: "text" },
      { key: "obs", label: "Observações", col: "Q", type: "textarea", span: 2 },
      { key: "dataSaida", label: "Data da saída", col: "T", type: "date" },
      { key: "destino", label: "Destino", col: "U", type: "text" },
      { key: "via", label: "Via", col: "V", type: "text" },
      { key: "entreguePor", label: "Entregue por", col: "W", type: "text" },
      { key: "recebidoPor", label: "Recebido por", col: "X", type: "text" },
      { key: "protocolo", label: "Protocolo de recebimento", col: "Y", type: "text" },
    ],
  },
};

/**
 * Listas de sugestão, lidas da aba oculta LISTAS em tempo de execução.
 * Mantém o app sincronizado com o que a planilha já usa.
 * Se um intervalo mudar na planilha, basta ajustar aqui.
 */
export const LISTS = {
  tipoDocEntrada: { range: "C46:C93", label: "Tipos de documento de entrada" },
  tipoDocSaida: { range: "B46:B84", label: "Documentos de saída e produção" },
  tipoManifestacao: { range: "B3:B10", label: "Tipo de manifestação" },
  natureza: { range: "C3:C38", label: "Natureza" },
  tipoProcessoCedmu: { range: "E3:E11", label: "Tipo de processo CEDMU" },
  unidade: { range: "F3:F18", label: "Unidades" },
  posto: { range: "H3:H15", label: "Posto / graduação" },
  ano: { range: "I3:I15", label: "Ano" },
  unidadeRecompensa: { range: "J3:J15", label: "Unidades de recompensa" },
  enquadramento: { range: "F20:F66", label: "Enquadramento disciplinar" },
  sancao: { range: "G20:G27", label: "Sanção" },
  recompensa: { range: "G32:G38", label: "Recompensa" },
  conclusao: { range: "B23:B33", label: "Conclusão" },
};

/** Colunas de conferência usadas na consulta (rótulos amigáveis). */
export const PREVIEW = {
  entrada: [
    { col: "A", label: "Prot." },
    { col: "C", label: "Data" },
    { col: "H", label: "Tipo" },
    { col: "I", label: "Nº doc" },
    { col: "L", label: "Origem" },
    { col: "M", label: "Remetente" },
    { col: "O", label: "Envolvido" },
  ],
  saida: [
    { col: "A", label: "Nº" },
    { col: "D", label: "Data" },
    { col: "I", label: "Tipo" },
    { col: "J", label: "Assunto" },
    { col: "K", label: "Envolvido" },
    { col: "U", label: "Destino" },
  ],
};

/** Colunas lidas na consulta (A..Y cobre as duas abas). */
export const CONSULTA_COLS = "A:Y";
