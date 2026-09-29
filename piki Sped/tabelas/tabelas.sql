# MySQL-Front Dump 2.5
#
# Host: localhost   Database: muitomais
# --------------------------------------------------------
# Server version 3.23.38-nt


#
# Table structure for table 'nfe_crt'
#

CREATE TABLE nfe_crt (
  id int(11) NOT NULL default '0',
  nome char(60) NOT NULL default '',
  PRIMARY KEY  (id)
) TYPE=MyISAM;



#
# Table structure for table 'nfe_cst'
#

CREATE TABLE nfe_cst (
  codigo char(2) NOT NULL default '',
  nome char(80) NOT NULL default '',
  PRIMARY KEY  (codigo)
) TYPE=MyISAM;



#
# Table structure for table 'tab_altpreco'
#

CREATE TABLE tab_altpreco (
  cod_pro char(8) NOT NULL default '',
  antigo_pro double(14,2) NOT NULL default '0.00',
  novo_pro double(14,2) NOT NULL default '0.00',
  data_alt datetime NOT NULL default '0000-00-00 00:00:00',
  cod_usu char(3) NOT NULL default '',
  status_alt char(1) NOT NULL default '1',
  PRIMARY KEY  (cod_pro,data_alt)
) TYPE=MyISAM;



#
# Table structure for table 'tab_aparencia'
#

CREATE TABLE tab_aparencia (
  cod_usu char(3) NOT NULL default '',
  CORFUNDO_PAR char(12) default NULL,
  CORCAIXA_PAR char(12) default NULL,
  CORFONTE_PAR char(12) default NULL,
  CORForm_PAR char(12) default NULL,
  CORFrame_PAR char(12) default NULL,
  CORFONTECAIXA_PAR char(12) default NULL,
  PATHPIC_PAR char(255) default NULL,
  TOOLVISIBLE_PAR char(1) default NULL,
  appearance_par char(1) default NULL,
  PRIMARY KEY  (cod_usu)
) TYPE=MyISAM;



#
# Table structure for table 'tab_baixapar'
#

CREATE TABLE tab_baixapar (
  seq_bai char(7) NOT NULL default '',
  seq_itd char(7) NOT NULL default '',
  valor_bai double(14,2) NOT NULL default '0.00',
  data_bai date NOT NULL default '0000-00-00',
  juros_bai double(14,2) NOT NULL default '0.00',
  desconto_bai double(14,2) NOT NULL default '0.00',
  tipo_bai char(3) NOT NULL default '',
  seq_doc char(7) NOT NULL default '',
  cod_usu char(3) default NULL,
  PRIMARY KEY  (seq_bai)
) TYPE=MyISAM;



#
# Table structure for table 'tab_bancos'
#

CREATE TABLE tab_bancos (
  COD_BCO char(3) NOT NULL default '',
  NOME_BCO char(50) NOT NULL default '',
  PRIMARY KEY  (COD_BCO)
) TYPE=MyISAM;



#
# Table structure for table 'tab_bolconfig'
#

CREATE TABLE tab_bolconfig (
  parcx_bol int(11) default NULL,
  parcy_bol int(11) default NULL,
  vencx_bol int(11) default NULL,
  vency_bol int(11) default NULL,
  emissx_bol int(11) default NULL,
  emissy_bol int(11) default NULL,
  numdocx_bol int(11) default NULL,
  numdocy_bol int(11) default NULL,
  espdocx_bol int(11) default NULL,
  espdocy_bol int(11) default NULL,
  aceitex_bol int(11) default NULL,
  aceitey_bol int(11) default NULL,
  dataprocx_bol int(11) default NULL,
  dataprocy_bol int(11) default NULL,
  instrx_bol int(11) default NULL,
  instry_bol int(11) default NULL,
  valorx_bol int(11) default NULL,
  valory_bol int(11) default NULL,
  sacadox_bol int(11) default NULL,
  sacadoy_bol int(11) default NULL,
  endx_bol int(11) default NULL,
  endy_bol int(11) default NULL,
  cnpjx_bol int(11) default NULL,
  cnpjy_bol int(11) unsigned default NULL,
  altura_bol int(11) unsigned default NULL,
  largura_bol int(11) unsigned default NULL,
  instr_bol char(100) default NULL,
  seq_bol int(10) unsigned default NULL,
  banco_bol char(30) default NULL
) TYPE=MyISAM;



#
# Table structure for table 'tab_boletos'
#

CREATE TABLE tab_boletos (
  seq_bol char(5) NOT NULL default '',
  seq_itd char(7) NOT NULL default '',
  valor_bol double(14,2) NOT NULL default '0.00',
  num_bol char(12) default NULL,
  venc_bol date NOT NULL default '0000-00-00',
  data_bol datetime NOT NULL default '0000-00-00 00:00:00',
  parc_bol char(5) NOT NULL default '',
  PRIMARY KEY  (seq_bol)
) TYPE=MyISAM;



#
# Table structure for table 'tab_caixa'
#

CREATE TABLE tab_caixa (
  SEQ_CAI char(7) NOT NULL default '',
  DATA_CAI date NOT NULL default '0000-00-00',
  VALOR_CAI double(14,2) NOT NULL default '0.00',
  entrada_CAI double(14,2) NOT NULL default '0.00',
  troco_CAI double(14,2) NOT NULL default '0.00',
  HISTORICO_CAI char(50) NOT NULL default '',
  TIPO_CAI char(1) NOT NULL default '',
  esp_cai char(3) NOT NULL default '',
  NUM_VEN char(7) default NULL,
  cod_con char(2) NOT NULL default '01',
  cod_cai char(2) NOT NULL default '01',
  ip_cai char(11) NOT NULL default '00000000000',
  cod_usu char(3) NOT NULL default '001',
  saldo_Cai double(14,2) NOT NULL default '0.00',
  cod_tur char(1) NOT NULL default '1',
  hora_cai char(8) NOT NULL default '00:00:00',
  PRIMARY KEY  (SEQ_CAI),
  KEY ffCodCai (cod_cai,cod_tur)
) TYPE=MyISAM;



#
# Table structure for table 'tab_caixas'
#

CREATE TABLE tab_caixas (
  cod_cai char(2) NOT NULL default '',
  nome_cai char(10) NOT NULL default '',
  ip_cai char(20) NOT NULL default '',
  COD_CON char(3) default NULL,
  status_Cai char(1) NOT NULL default 'A',
  cod_tur char(1) NOT NULL default '1',
  impressora_cai char(1) NOT NULL default 'B',
  modo_cai char(1) NOT NULL default '0',
  bloqueado_cai char(1) NOT NULL default 'N',
  PRIMARY KEY  (cod_cai)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cancelamentos'
#

CREATE TABLE tab_cancelamentos (
  num_ven char(7) NOT NULL default '',
  data_can datetime NOT NULL default '0000-00-00 00:00:00',
  valor_ven double(14,2) NOT NULL default '0.00',
  tipo_ven char(1) NOT NULL default 'R',
  cod_cai char(2) NOT NULL default '',
  cod_tur char(2) NOT NULL default '',
  PRIMARY KEY  (num_ven)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cargos'
#

CREATE TABLE tab_cargos (
  COD_CAR char(3) NOT NULL default '',
  NOME_CAR char(20) NOT NULL default '',
  venda_Car char(1) NOT NULL default 'N',
  PRIMARY KEY  (COD_CAR)
) TYPE=MyISAM;



#
# Table structure for table 'tab_categorias'
#

CREATE TABLE tab_categorias (
  cod_cat char(3) NOT NULL default '',
  nome_cat char(40) NOT NULL default '',
  PRIMARY KEY  (cod_cat)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cfop'
#

CREATE TABLE tab_cfop (
  cfop_cfo char(4) NOT NULL default '',
  nome_cfo char(40) NOT NULL default '',
  tipo_cfo char(1) NOT NULL default 'S',
  PRIMARY KEY  (cfop_cfo)
) TYPE=MyISAM;



#
# Table structure for table 'tab_chequedep'
#

CREATE TABLE tab_chequedep (
  seq_che char(7) NOT NULL default '',
  data_dep date NOT NULL default '0000-00-00',
  cod_con char(3) NOT NULL default '',
  PRIMARY KEY  (seq_che,data_dep)
) TYPE=MyISAM;



#
# Table structure for table 'tab_chequedev'
#

CREATE TABLE tab_chequedev (
  seq_che char(7) NOT NULL default '',
  seq_dev int(11) NOT NULL default '0',
  data_dev date NOT NULL default '0000-00-00',
  motivo_dev char(30) NOT NULL default '',
  PRIMARY KEY  (seq_che,seq_dev)
) TYPE=MyISAM;



#
# Table structure for table 'tab_chequepag'
#

CREATE TABLE tab_chequepag (
  seq_che char(7) NOT NULL default '',
  num_che char(12) NOT NULL default '',
  cod_con char(3) NOT NULL default '',
  PRIMARY KEY  (seq_che)
) TYPE=MyISAM;



#
# Table structure for table 'tab_chequerec'
#

CREATE TABLE tab_chequerec (
  seq_che char(7) NOT NULL default '',
  num_che char(12) NOT NULL default '',
  cod_bco char(3) NOT NULL default '',
  age_che char(8) NOT NULL default '',
  cco_che char(12) NOT NULL default '',
  cpf_che char(18) NOT NULL default '',
  nome_che char(50) NOT NULL default '',
  PRIMARY KEY  (seq_che)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cheques1'
#

CREATE TABLE tab_cheques1 (
  seq_che char(7) NOT NULL default '',
  num_che char(12) NOT NULL default '',
  cod_bco char(3) NOT NULL default '',
  age_che char(8) default NULL,
  cco_che char(12) default NULL,
  nome_che char(50) default NULL,
  cpf_che char(18) NOT NULL default '',
  valor_che double(14,2) default NULL,
  venc_che date NOT NULL default '0000-00-00',
  cod_ter char(5) NOT NULL default '',
  tipo_che char(1) NOT NULL default '',
  seq_doc char(7) default NULL,
  parc_che int(11) NOT NULL default '0',
  num_ven char(7) default NULL,
  dep_che date NOT NULL default '0000-00-00',
  cod_usu char(3) default NULL,
  cod_con char(3) NOT NULL default '001',
  status_Che char(1) NOT NULL default '1',
  cod_loc char(1) NOT NULL default '1',
  hora_che char(5) NOT NULL default '00:00',
  PRIMARY KEY  (seq_che)
) TYPE=MyISAM;



#
# Table structure for table 'tab_chequesb'
#

CREATE TABLE tab_chequesb (
  seq_che char(7) NOT NULL default '',
  num_che char(12) NOT NULL default '',
  cod_bco char(3) NOT NULL default '',
  age_che char(8) default NULL,
  cco_che char(12) default NULL,
  nome_che char(50) default NULL,
  cpf_che char(18) NOT NULL default '',
  valor_che double(14,2) default NULL,
  venc_che date NOT NULL default '0000-00-00',
  cod_ter char(5) NOT NULL default '',
  tipo_che char(1) NOT NULL default '',
  seq_doc char(7) default NULL,
  parc_che int(11) NOT NULL default '0',
  num_ven char(7) default NULL,
  dep_che date NOT NULL default '0000-00-00',
  cod_usu char(3) default NULL,
  cod_con char(3) NOT NULL default '001',
  status_Che char(1) NOT NULL default '1',
  PRIMARY KEY  (seq_che)
) TYPE=MyISAM;



#
# Table structure for table 'tab_chequesback'
#

CREATE TABLE tab_chequesback (
  seq_che char(7) NOT NULL default '',
  num_che char(12) NOT NULL default '',
  cod_bco char(3) NOT NULL default '',
  age_che char(8) default NULL,
  cco_che char(12) default NULL,
  nome_che char(50) default NULL,
  cpf_che char(18) NOT NULL default '',
  valor_che double(14,2) default NULL,
  venc_che date NOT NULL default '0000-00-00',
  cod_ter char(5) NOT NULL default '',
  tipo_che char(1) NOT NULL default '',
  seq_doc char(7) default NULL,
  parc_che int(11) NOT NULL default '0',
  num_ven char(7) default NULL,
  dep_che date NOT NULL default '0000-00-00',
  cod_usu char(3) default NULL,
  cod_con char(3) NOT NULL default '001',
  status_Che char(1) NOT NULL default '1',
  PRIMARY KEY  (seq_che)
) TYPE=MyISAM;



#
# Table structure for table 'tab_chv'
#

CREATE TABLE tab_chv (
  seq_cai char(7) NOT NULL default '',
  num_che char(12) NOT NULL default '',
  cod_bco char(3) NOT NULL default '',
  age_che char(8) NOT NULL default '',
  cco_che char(12) NOT NULL default '',
  cpf_che char(18) NOT NULL default '',
  nome_che char(50) NOT NULL default '',
  valor_che double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cai)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cidades'
#

CREATE TABLE tab_cidades (
  cod_cid varchar(7) NOT NULL default '',
  nome_cid varchar(60) NOT NULL default '',
  estado_cid varchar(30) NOT NULL default '',
  uf_cid char(2) NOT NULL default '',
  PRIMARY KEY  (cod_cid)
) TYPE=MyISAM;



#
# Table structure for table 'tab_class'
#

CREATE TABLE tab_class (
  seq_fin char(7) NOT NULL default '',
  cod_pla char(3) NOT NULL default '',
  valor_cla double(14,2) NOT NULL default '0.00',
  cod_usu char(3) NOT NULL default '',
  PRIMARY KEY  (seq_fin,cod_pla)
) TYPE=MyISAM;



#
# Table structure for table 'tab_clientes'
#

CREATE TABLE tab_clientes (
  COD_CLI char(5) NOT NULL default '',
  NOME_CLI char(50) default NULL,
  CPF_CLI char(18) default NULL,
  RG_CLI char(16) default NULL,
  ENDERECO_CLI char(30) default NULL,
  ENDFIRMA_CLI char(30) default NULL,
  BAIRRO_CLI char(20) default NULL,
  BAIRROFIRMA_CLI char(20) default NULL,
  CIDADE_CLI char(30) default NULL,
  UF_CLI char(2) default NULL,
  CEP_CLI char(12) default NULL,
  TELEFONE_CLI char(20) default NULL,
  fax_cli char(20) default NULL,
  celular_cli char(20) default NULL,
  TIPO_CLI char(1) default 'F',
  NASC_CLI date default NULL,
  ESTADOCIV_CLI char(1) default NULL,
  PROFISSAO_CLI char(30) default NULL,
  MAE_CLI char(50) default NULL,
  REF1_CLI char(50) default NULL,
  REF2_CLI char(50) default NULL,
  STATUS_CLI char(1) default '1',
  REFPN1_CLI char(50) default NULL,
  REFPE1_CLI char(50) default NULL,
  REFPT1_CLI char(50) default NULL,
  REFPN2_CLI char(50) default NULL,
  REFPE2_CLI char(50) default NULL,
  REFPT2_CLI char(50) default NULL,
  FIRMA_CLI char(30) default NULL,
  CARGO_CLI char(20) default NULL,
  TELFIRMA_CLI char(14) default NULL,
  CIDFIRMA_CLI char(30) default NULL,
  UFFIRMA_CLI char(2) default NULL,
  RENDA_CLI double(14,2) NOT NULL default '0.00',
  obs_cli char(240) default NULL,
  cel_cli char(20) default NULL,
  LIMITE_CLI double(14,2) NOT NULL default '0.00',
  dataalt_cli date default '0000-00-00',
  inscprodrur_cli char(12) default NULL,
  cod_cid char(7) default '5218300',
  complemento_cli char(40) default NULL,
  Numero_cli char(10) default NULL,
  ie_cli char(12) default NULL,
  PRIMARY KEY  (COD_CLI)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cmp'
#

CREATE TABLE tab_cmp (
  COD_FOR char(5) NOT NULL default '',
  NUM_NOT char(16) NOT NULL default '',
  SERIE_NOT char(4) NOT NULL default '',
  VALOR_NOT double(14,2) NOT NULL default '0.00',
  BC_NOT double(14,2) default NULL,
  FRETE_NOT double(14,2) default NULL,
  ICMS_NOT double(14,2) default NULL,
  IPI_NOT double(14,2) default NULL,
  DESPESAS_NOT double(14,2) default NULL,
  OPERACAO_NOT char(6) NOT NULL default '',
  EMISSAO_NOT date NOT NULL default '0000-00-00',
  ENTRADA_NOT date NOT NULL default '0000-00-00',
  COD_USU char(3) NOT NULL default '',
  STATUS_NOT char(1) NOT NULL default '',
  valorprod_not double(14,2) NOT NULL default '0.00',
  seq_not char(7) NOT NULL default '0000001',
  desconto_not double(14,2) NOT NULL default '0.00',
  data_not datetime NOT NULL default '0000-00-00 00:00:00',
  seq_ent char(7) NOT NULL default '0000001',
  PRIMARY KEY  (seq_not),
  KEY codfor (COD_FOR)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cmptmp'
#

CREATE TABLE tab_cmptmp (
  COD_FOR char(5) NOT NULL default '',
  NUM_NOT char(16) NOT NULL default '',
  SERIE_NOT char(4) NOT NULL default '',
  VALOR_NOT double(14,2) NOT NULL default '0.00',
  BC_NOT double(14,2) default NULL,
  FRETE_NOT double(14,2) default NULL,
  ICMS_NOT double(14,2) default NULL,
  IPI_NOT double(14,2) default NULL,
  DESPESAS_NOT double(14,2) default NULL,
  OPERACAO_NOT char(6) NOT NULL default '',
  EMISSAO_NOT date NOT NULL default '0000-00-00',
  ENTRADA_NOT date NOT NULL default '0000-00-00',
  COD_USU char(3) NOT NULL default '',
  STATUS_NOT char(1) NOT NULL default '',
  valorprod_not double(14,2) NOT NULL default '0.00',
  seq_not char(7) NOT NULL default '0000001',
  desconto_not double(14,2) NOT NULL default '0.00',
  data_not datetime NOT NULL default '0000-00-00 00:00:00',
  seq_ent char(7) NOT NULL default '0000001',
  PRIMARY KEY  (seq_not)
) TYPE=MyISAM;



#
# Table structure for table 'tab_codbarras'
#

CREATE TABLE tab_codbarras (
  cod_pro char(8) NOT NULL default '',
  codbarra_pro char(16) NOT NULL default '',
  PRIMARY KEY  (cod_pro,codbarra_pro)
) TYPE=MyISAM;



#
# Table structure for table 'tab_codigos'
#

CREATE TABLE tab_codigos (
  table_cod char(30) NOT NULL default '',
  field_cod char(30) NOT NULL default '',
  value_cod char(15) NOT NULL default '',
  PRIMARY KEY  (table_cod,field_cod)
) TYPE=MyISAM;



#
# Table structure for table 'tab_compras'
#

CREATE TABLE tab_compras (
  COD_FOR char(5) NOT NULL default '',
  NUM_NOT char(16) NOT NULL default '',
  SERIE_NOT char(4) NOT NULL default '',
  VALOR_NOT double(14,2) NOT NULL default '0.00',
  BC_NOT double(14,2) default NULL,
  FRETE_NOT double(14,2) default NULL,
  ICMS_NOT double(14,2) default NULL,
  IPI_NOT double(14,2) default NULL,
  DESPESAS_NOT double(14,2) default NULL,
  OPERACAO_NOT char(6) NOT NULL default '',
  EMISSAO_NOT date NOT NULL default '0000-00-00',
  ENTRADA_NOT date NOT NULL default '0000-00-00',
  COD_USU char(3) NOT NULL default '',
  STATUS_NOT char(1) NOT NULL default '',
  valorprod_not double(14,2) NOT NULL default '0.00',
  seq_not char(7) NOT NULL default '0000001',
  desconto_not double(14,2) NOT NULL default '0.00',
  data_not datetime NOT NULL default '0000-00-00 00:00:00',
  PRIMARY KEY  (seq_not)
) TYPE=MyISAM;



#
# Table structure for table 'tab_comprastmp'
#

CREATE TABLE tab_comprastmp (
  COD_FOR char(5) NOT NULL default '',
  NUM_NOT char(16) NOT NULL default '',
  SERIE_NOT char(4) NOT NULL default '',
  VALOR_NOT double(14,2) NOT NULL default '0.00',
  BC_NOT double(14,2) default NULL,
  FRETE_NOT double(14,2) default NULL,
  ICMS_NOT double(14,2) default NULL,
  IPI_NOT double(14,2) default NULL,
  DESPESAS_NOT double(14,2) default NULL,
  OPERACAO_NOT char(6) NOT NULL default '',
  EMISSAO_NOT date NOT NULL default '0000-00-00',
  ENTRADA_NOT date NOT NULL default '0000-00-00',
  COD_USU char(3) NOT NULL default '',
  STATUS_NOT char(1) NOT NULL default '',
  valorprod_not double(14,2) NOT NULL default '0.00',
  seq_not char(7) NOT NULL default '0000001',
  desconto_not double(14,2) NOT NULL default '0.00',
  data_not datetime NOT NULL default '0000-00-00 00:00:00',
  PRIMARY KEY  (seq_not)
) TYPE=MyISAM;



#
# Table structure for table 'tab_confignf'
#

CREATE TABLE tab_confignf (
  natopx_con int(11) NOT NULL default '0',
  natopy_con int(11) NOT NULL default '0',
  cfopx_con int(11) NOT NULL default '0',
  cfopy_con int(11) NOT NULL default '0',
  iex_con int(11) NOT NULL default '0',
  iey_con int(11) NOT NULL default '0',
  nomedestx_con int(11) NOT NULL default '0',
  nomedesty_con int(11) NOT NULL default '0',
  cnpjx_con int(11) NOT NULL default '0',
  cnpjy_con int(11) NOT NULL default '0',
  emissx_con int(11) NOT NULL default '0',
  emissy_con int(11) NOT NULL default '0',
  enddestx_con int(11) NOT NULL default '0',
  enddesty_con int(11) NOT NULL default '0',
  baidestx_con int(11) NOT NULL default '0',
  baidesty_con int(11) NOT NULL default '0',
  cepdestx_con int(11) NOT NULL default '0',
  cepdesty_con int(11) NOT NULL default '0',
  munidestx_con int(11) NOT NULL default '0',
  munidesty_con int(11) NOT NULL default '0',
  fonedestx_con int(11) NOT NULL default '0',
  fonedesty_con int(11) NOT NULL default '0',
  ufdestx_con int(11) NOT NULL default '0',
  ufdesty_con int(11) NOT NULL default '0',
  iedestx_con int(11) NOT NULL default '0',
  iedesty_con int(11) NOT NULL default '0',
  iniprody_con int(11) NOT NULL default '0',
  codprod_con int(11) NOT NULL default '0',
  descprod_con int(11) NOT NULL default '0',
  stprod_con int(11) NOT NULL default '0',
  unid_con int(11) NOT NULL default '0',
  qtdprod_con int(11) NOT NULL default '0',
  unitprod_con int(11) NOT NULL default '0',
  totalprod_con int(11) NOT NULL default '0',
  aliqprod_con int(11) NOT NULL default '0',
  distprod_con int(11) NOT NULL default '0',
  bcx_con int(11) NOT NULL default '0',
  bcy_con int(11) NOT NULL default '0',
  vlricmsx_con int(11) NOT NULL default '0',
  vlricmsy_con int(11) NOT NULL default '0',
  bcsubsx_con int(11) NOT NULL default '0',
  bcsubsy_con int(11) NOT NULL default '0',
  vlrsubsx_con int(11) NOT NULL default '0',
  vlrsubsy_con int(11) NOT NULL default '0',
  vlrprodx_con int(11) NOT NULL default '0',
  vlrprody_con int(11) NOT NULL default '0',
  vlrfretex_con int(11) NOT NULL default '0',
  vlrfretey_con int(11) NOT NULL default '0',
  vlrsegx_con int(11) NOT NULL default '0',
  vlrsegy_con int(11) NOT NULL default '0',
  outrasdespx_con int(11) NOT NULL default '0',
  outrasdespy_con int(11) NOT NULL default '0',
  vlripix_con int(11) NOT NULL default '0',
  vlripiy_con int(11) NOT NULL default '0',
  vlrtotalx_con int(11) NOT NULL default '0',
  vlrtotaly_con int(11) NOT NULL default '0',
  nomefrex_con int(11) NOT NULL default '0',
  nomefrey_con int(11) NOT NULL default '0',
  tipofrex_con int(11) NOT NULL default '0',
  tipofrey_con int(11) NOT NULL default '0',
  placafrex_con int(11) NOT NULL default '0',
  placafrey_con int(11) NOT NULL default '0',
  ufveix_con int(11) NOT NULL default '0',
  ufveiy_con int(11) NOT NULL default '0',
  cnpjfrex_con int(11) NOT NULL default '0',
  cnpjfrey_con int(11) NOT NULL default '0',
  endfrex_con int(11) NOT NULL default '0',
  endfrey_con int(11) NOT NULL default '0',
  munifrex_con int(11) NOT NULL default '0',
  munifrey_con int(11) NOT NULL default '0',
  uffrex_con int(11) NOT NULL default '0',
  uffrey_con int(11) NOT NULL default '0',
  iefrex_con int(11) NOT NULL default '0',
  iefrey_con int(11) NOT NULL default '0',
  qtdfrex_con int(11) NOT NULL default '0',
  qtdfrey_con int(11) NOT NULL default '0',
  espfrex_con int(11) NOT NULL default '0',
  espfrey_con int(11) NOT NULL default '0',
  marcafrex_con int(11) NOT NULL default '0',
  marcafrey_con int(11) NOT NULL default '0',
  numfrex_con int(11) NOT NULL default '0',
  numfrey_con int(11) NOT NULL default '0',
  pbrutofrex_con int(11) NOT NULL default '0',
  pbrutofrey_con int(11) NOT NULL default '0',
  pliqfrex_con int(11) NOT NULL default '0',
  pliqfrey_con int(11) NOT NULL default '0',
  dadosaddx_con tinyint(3) unsigned default NULL,
  dadosaddy_con tinyint(3) unsigned default NULL,
  numrodapex_con tinyint(3) unsigned default NULL,
  numrodapey_con int(11) NOT NULL default '0',
  numx_con int(11) NOT NULL default '0',
  numy_con int(11) NOT NULL default '0',
  marcax_con int(11) NOT NULL default '0',
  marcay_con int(11) NOT NULL default '0'
) TYPE=MyISAM;



#
# Table structure for table 'tab_contagemdireta'
#

CREATE TABLE tab_contagemdireta (
  id int(10) NOT NULL auto_increment,
  codigo char(8) NOT NULL default '',
  qtd_anterior double(14,2) NOT NULL default '0.00',
  qtd_nova double(14,2) NOT NULL default '0.00',
  cod_usu char(3) NOT NULL default '',
  data datetime default NULL,
  PRIMARY KEY  (id)
) TYPE=MyISAM;



#
# Table structure for table 'tab_contagens'
#

CREATE TABLE tab_contagens (
  seq_cnt char(4) NOT NULL default '',
  data_cnt datetime NOT NULL default '0000-00-00 00:00:00',
  cod_usu char(3) NOT NULL default '',
  status_cnt char(1) NOT NULL default 'N',
  tipo_con char(1) NOT NULL default 'R',
  PRIMARY KEY  (seq_cnt)
) TYPE=MyISAM;



#
# Table structure for table 'tab_contas'
#

CREATE TABLE tab_contas (
  cod_con char(3) NOT NULL default '',
  nome_con char(20) NOT NULL default '',
  num_con char(12) default NULL,
  cod_bco char(3) default NULL,
  age_con char(12) default NULL,
  tit_con char(30) default NULL,
  status_con char(1) NOT NULL default 'H',
  PRIMARY KEY  (cod_con)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cupons'
#

CREATE TABLE tab_cupons (
  seq_cup int(11) NOT NULL default '0',
  num_cup char(8) NOT NULL default '',
  num_ven char(7) NOT NULL default '',
  valor_cup double(14,2) NOT NULL default '0.00',
  desconto_cup double(14,2) NOT NULL default '0.00',
  data_cup datetime NOT NULL default '0000-00-00 00:00:00',
  cod_cai char(2) NOT NULL default '',
  cod_tur char(1) NOT NULL default '',
  cancelado_cup char(1) NOT NULL default 'N',
  icms_cup double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cup),
  KEY fVen (num_ven)
) TYPE=MyISAM;



#
# Table structure for table 'tab_cuponstmp'
#

CREATE TABLE tab_cuponstmp (
  seq_cup int(11) NOT NULL default '0',
  num_cup char(8) NOT NULL default '',
  num_ven char(7) NOT NULL default '',
  valor_cup double(14,2) NOT NULL default '0.00',
  desconto_cup double(14,2) NOT NULL default '0.00',
  data_cup datetime NOT NULL default '0000-00-00 00:00:00',
  cod_cai char(2) NOT NULL default '',
  cod_tur char(1) NOT NULL default '',
  cancelado_cup char(1) NOT NULL default 'N',
  icms_cup double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cup)
) TYPE=MyISAM;



#
# Table structure for table 'tab_defusu'
#

CREATE TABLE tab_defusu (
  cod_usu char(3) NOT NULL default '',
  abrevenda_usu char(1) NOT NULL default 'N',
  abrenota_usu char(1) NOT NULL default 'N',
  efetnota_usu char(1) NOT NULL default 'N',
  abresaida_usu char(1) NOT NULL default 'N',
  lancaixa_usu char(1) NOT NULL default 'N',
  landoc_usu char(1) NOT NULL default 'N',
  lanche_usu char(1) NOT NULL default 'N',
  baixadoc_usu char(1) NOT NULL default 'N',
  baixache_usu char(1) NOT NULL default 'N',
  edicai_usu char(1) NOT NULL default 'N',
  exclcai_usu char(1) NOT NULL default 'N',
  excldoc_usu char(1) NOT NULL default 'N',
  cadprod_usu char(1) NOT NULL default 'N',
  altprod_usu char(1) NOT NULL default 'N',
  altpreco_usu char(1) NOT NULL default 'N',
  cadusu_usu char(1) NOT NULL default 'N',
  cadcli_usu char(1) NOT NULL default 'N',
  cadcco_usu char(1) NOT NULL default 'N',
  cadformaspag_usu char(1) NOT NULL default 'N',
  altformaspag_usu char(1) NOT NULL default 'N',
  viscusto_usu char(1) NOT NULL default 'N',
  visrelest_usu char(1) NOT NULL default 'N',
  visrelfin_usu char(1) NOT NULL default 'N',
  fechacai_usu char(1) NOT NULL default 'N',
  fechames_usu char(1) NOT NULL default 'N',
  desconto_usu double(14,2) NOT NULL default '0.00',
  contrpre_usu char(1) NOT NULL default 'N',
  exclvenda_usu char(1) NOT NULL default 'N',
  PRIMARY KEY  (cod_usu)
) TYPE=MyISAM;



#
# Table structure for table 'tab_devolucoes'
#

CREATE TABLE tab_devolucoes (
  num_ven char(7) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) default NULL,
  qtd_dev double(14,2) default NULL,
  data_dev datetime default NULL,
  gerado_dev double(14,2) default NULL,
  cod_usu char(3) NOT NULL default '',
  tipo_dev char(1) NOT NULL default 'I',
  SEQEFET_DEV char(7) default NULL,
  cod_cli char(5) default NULL,
  desconto_dev double(14,2) default '0.00',
  cod_tip char(2) default NULL,
  situacao_dev char(1) default NULL,
  PRIMARY KEY  (num_ven,cod_pro,tipo_dev)
) TYPE=MyISAM;



#
# Table structure for table 'tab_doc1'
#

CREATE TABLE tab_doc1 (
  seq_doc char(7) NOT NULL default '',
  num_doc char(12) NOT NULL default '',
  valor_doc double(14,2) default NULL,
  data_doc date NOT NULL default '0000-00-00',
  cod_ter char(5) NOT NULL default '',
  tipo_doc char(1) NOT NULL default '',
  num_ven char(7) default NULL,
  parc_doc int(11) default NULL,
  historico_doc char(40) default NULL,
  cod_usu char(3) default NULL,
  excluido_doc char(1) NOT NULL default 'N',
  cod_pla char(3) NOT NULL default '000',
  esp_doc char(1) NOT NULL default 'V',
  PRIMARY KEY  (seq_doc)
) TYPE=MyISAM;



#
# Table structure for table 'tab_docback'
#

CREATE TABLE tab_docback (
  seq_doc char(7) NOT NULL default '',
  num_doc char(12) NOT NULL default '',
  valor_doc double(14,2) default NULL,
  data_doc date NOT NULL default '0000-00-00',
  cod_ter char(5) NOT NULL default '',
  tipo_doc char(1) NOT NULL default '',
  num_ven char(7) default NULL,
  parc_doc int(11) default NULL,
  historico_doc char(40) default NULL,
  cod_usu char(3) default NULL,
  excluido_doc char(1) NOT NULL default 'N',
  PRIMARY KEY  (seq_doc)
) TYPE=MyISAM;



#
# Table structure for table 'tab_ecfimp'
#

CREATE TABLE tab_ecfimp (
  cod_imp char(3) NOT NULL default '',
  ecf_imp char(3) NOT NULL default '',
  impressora_imp char(1) NOT NULL default '',
  PRIMARY KEY  (cod_imp,impressora_imp)
) TYPE=MyISAM;



#
# Table structure for table 'tab_embalagens'
#

CREATE TABLE tab_embalagens (
  cod_pro char(8) NOT NULL default '',
  qtd_emb int(11) NOT NULL default '0',
  PRIMARY KEY  (cod_pro,qtd_emb)
) TYPE=MyISAM;



#
# Table structure for table 'tab_fechacaixa'
#

CREATE TABLE tab_fechacaixa (
  DATA_CAI date NOT NULL default '0000-00-00',
  VALOR_FEC double(14,2) NOT NULL default '0.00',
  COD_USU char(10) NOT NULL default '',
  cod_cai char(2) NOT NULL default '',
  fundo_cai double(14,2) NOT NULL default '0.00',
  turno_cai char(1) NOT NULL default '1',
  PRIMARY KEY  (DATA_CAI,cod_cai)
) TYPE=MyISAM;



#
# Table structure for table 'tab_financeiras'
#

CREATE TABLE tab_financeiras (
  COD_FIN char(2) NOT NULL default '',
  COD_TIP char(2) NOT NULL default '',
  TIPO_FIN char(1) NOT NULL default '',
  NOME_FIN char(20) NOT NULL default '',
  TEL_FIN char(30) default NULL,
  PRIMARY KEY  (COD_FIN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_financeiro'
#

CREATE TABLE tab_financeiro (
  SEQ_FIN char(7) NOT NULL default '',
  DATA_FIN date NOT NULL default '0000-00-00',
  VALOR_FIN double(14,2) NOT NULL default '0.00',
  HISTORICO_FIN char(50) NOT NULL default '',
  TIPO_FIN char(1) NOT NULL default '',
  ESP_FIN char(3) NOT NULL default '',
  cod_con char(3) NOT NULL default '01',
  saldo_fin double(14,2) default NULL,
  seq_lan char(7) default NULL,
  num_ven char(7) default NULL,
  cod_usu char(3) default NULL,
  cod_pla char(3) NOT NULL default '000',
  PRIMARY KEY  (SEQ_FIN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_financeiroback'
#

CREATE TABLE tab_financeiroback (
  SEQ_FIN char(7) NOT NULL default '',
  DATA_FIN date NOT NULL default '0000-00-00',
  VALOR_FIN double(14,2) NOT NULL default '0.00',
  HISTORICO_FIN char(50) NOT NULL default '',
  TIPO_FIN char(1) NOT NULL default '',
  ESP_FIN char(3) NOT NULL default '',
  cod_con char(3) NOT NULL default '01',
  saldo_fin double(14,2) default NULL,
  seq_lan char(7) default NULL,
  num_ven char(7) default NULL,
  cod_usu char(3) default NULL,
  PRIMARY KEY  (SEQ_FIN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_fornecedores'
#

CREATE TABLE tab_fornecedores (
  COD_FOR char(5) NOT NULL default '',
  NOME_FOR char(50) NOT NULL default '',
  FANTASIA_FOR char(50) default NULL,
  CNPJ_FOR char(18) NOT NULL default '',
  IE_FOR char(12) default NULL,
  ENDERECO_FOR char(30) default NULL,
  BAIRRO_FOR char(20) default NULL,
  CIDADE_FOR char(30) default NULL,
  UF_FOR char(2) default NULL,
  CEP_FOR char(8) default NULL,
  TELEFONE_FOR char(16) default NULL,
  FAX_FOR char(16) default NULL,
  CONTATO_FOR char(30) default NULL,
  EMAIL_FOR char(64) default NULL,
  codantigo_for char(5) default NULL,
  obs_for char(60) default NULL,
  visivel_for char(1) NOT NULL default 'S',
  PRIMARY KEY  (COD_FOR)
) TYPE=MyISAM;



#
# Table structure for table 'tab_fornpcp'
#

CREATE TABLE tab_fornpcp (
  seq_pcp char(6) NOT NULL default '',
  cod_Ter char(6) NOT NULL default '',
  sts_fpc char(1) NOT NULL default '',
  PRIMARY KEY  (seq_pcp,cod_Ter)
) TYPE=MyISAM;



#
# Table structure for table 'tab_fotos'
#

CREATE TABLE tab_fotos (
  seq_fot varchar(7) NOT NULL default '',
  cod_pro varchar(8) NOT NULL default '',
  foto_fot blob,
  PRIMARY KEY  (seq_fot)
) TYPE=MyISAM;



#
# Table structure for table 'tab_funcionarios'
#

CREATE TABLE tab_funcionarios (
  COD_FUN char(3) NOT NULL default '',
  NOME_FUN char(50) NOT NULL default '',
  GUERRA_FUN char(16) NOT NULL default '',
  COD_CAR char(3) default NULL,
  RG_FUN char(18) default NULL,
  CPF_FUN char(18) default NULL,
  ENDERECO_FUN char(64) default NULL,
  TELEFONE_FUN char(16) default NULL,
  STATUS_FUN char(1) default NULL,
  PRIMARY KEY  (COD_FUN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_imob'
#

CREATE TABLE tab_imob (
  cod_imo char(5) NOT NULL default '',
  unidade_imo char(3) NOT NULL default '',
  desc_imo char(50) NOT NULL default '',
  qtd_imo double(14,2) NOT NULL default '0.00',
  unit_imo double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (cod_imo)
) TYPE=MyISAM;



#
# Table structure for table 'tab_impostos'
#

CREATE TABLE tab_impostos (
  cod_imp char(3) NOT NULL default '',
  desc_imp char(14) NOT NULL default '',
  perc_imp double(14,2) NOT NULL default '0.00',
  ecf_imp char(3) NOT NULL default '',
  PRIMARY KEY  (cod_imp)
) TYPE=MyISAM;



#
# Table structure for table 'tab_invent'
#

CREATE TABLE tab_invent (
  cod_pro char(8) NOT NULL default '',
  nome_pro char(50) NOT NULL default '',
  unidade_pro char(3) NOT NULL default '',
  custo_pro double(14,2) NOT NULL default '0.00',
  customedio_pro double(14,2) NOT NULL default '0.00',
  preco_pro double(14,2) NOT NULL default '0.00',
  estoque_pro double(14,2) NOT NULL default '0.00',
  reservado_pro double(14,2) NOT NULL default '0.00',
  datahora_inv datetime NOT NULL default '0000-00-00 00:00:00',
  ef_pro double(14,2) NOT NULL default '0.00',
  cf_pro double(14,2) NOT NULL default '0.00',
  KEY cod_pro (cod_pro,nome_pro,datahora_inv)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcmp'
#

CREATE TABLE tab_itcmp (
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,4) NOT NULL default '0.0000',
  QTD_ITE double(14,2) NOT NULL default '0.00',
  IPI_ITE double(14,4) default NULL,
  GASTOS_ITE double(14,4) default NULL,
  CUSTO_ITE double(14,4) default NULL,
  SEQ_ITE int(11) NOT NULL default '0',
  sujerido_ite double(14,2) NOT NULL default '0.00',
  precofuturo_ite double(14,2) NOT NULL default '0.00',
  lucrosuj_ite double(14,2) NOT NULL default '0.00',
  lucroliq_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,4) NOT NULL default '0.0000',
  seq_not char(7) NOT NULL default '',
  custosai_ite double(14,4) NOT NULL default '0.0000',
  icms_ite double(14,2) NOT NULL default '0.00',
  qtd_emb int(10) unsigned NOT NULL default '1',
  PRIMARY KEY  (seq_not,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcmptmp'
#

CREATE TABLE tab_itcmptmp (
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,4) NOT NULL default '0.0000',
  QTD_ITE double(14,2) NOT NULL default '0.00',
  IPI_ITE double(14,4) default NULL,
  GASTOS_ITE double(14,4) default NULL,
  CUSTO_ITE double(14,4) default NULL,
  SEQ_ITE int(11) NOT NULL default '0',
  sujerido_ite double(14,2) NOT NULL default '0.00',
  precofuturo_ite double(14,2) NOT NULL default '0.00',
  lucrosuj_ite double(14,2) NOT NULL default '0.00',
  lucroliq_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,4) NOT NULL default '0.0000',
  seq_not char(7) NOT NULL default '',
  custosai_ite double(14,4) NOT NULL default '0.0000',
  icms_ite double(14,2) NOT NULL default '0.00',
  qtd_emb int(10) unsigned NOT NULL default '1',
  PRIMARY KEY  (seq_not,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcompras'
#

CREATE TABLE tab_itcompras (
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,4) NOT NULL default '0.0000',
  QTD_ITE double(14,2) NOT NULL default '0.00',
  IPI_ITE double(14,4) default NULL,
  GASTOS_ITE double(14,4) default NULL,
  CUSTO_ITE double(14,4) default NULL,
  SEQ_ITE int(11) NOT NULL default '0',
  sujerido_ite double(14,2) NOT NULL default '0.00',
  precofuturo_ite double(14,2) NOT NULL default '0.00',
  lucrosuj_ite double(14,2) NOT NULL default '0.00',
  lucroliq_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,4) NOT NULL default '0.0000',
  seq_not char(7) NOT NULL default '',
  custosai_ite double(14,4) NOT NULL default '0.0000',
  icms_ite double(14,2) NOT NULL default '0.00',
  qtd_emb int(10) unsigned NOT NULL default '1',
  PRIMARY KEY  (seq_not,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcomprastmp'
#

CREATE TABLE tab_itcomprastmp (
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,4) NOT NULL default '0.0000',
  QTD_ITE double(14,2) NOT NULL default '0.00',
  IPI_ITE double(14,4) default NULL,
  GASTOS_ITE double(14,4) default NULL,
  CUSTO_ITE double(14,4) default NULL,
  SEQ_ITE int(11) NOT NULL default '0',
  sujerido_ite double(14,2) NOT NULL default '0.00',
  precofuturo_ite double(14,2) NOT NULL default '0.00',
  lucrosuj_ite double(14,2) NOT NULL default '0.00',
  lucroliq_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,4) NOT NULL default '0.0000',
  seq_not char(7) NOT NULL default '',
  custosai_ite double(14,4) NOT NULL default '0.0000',
  icms_ite double(14,2) NOT NULL default '0.00',
  qtd_emb int(10) unsigned NOT NULL default '1',
  PRIMARY KEY  (seq_not,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcont'
#

CREATE TABLE tab_itcont (
  seq_cnt char(4) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_itc double(14,2) NOT NULL default '0.00',
  data_itc datetime NOT NULL default '0000-00-00 00:00:00',
  PRIMARY KEY  (seq_cnt,cod_pro)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcup'
#

CREATE TABLE tab_itcup (
  seq_cup int(11) NOT NULL default '0',
  seq_ite char(8) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,2) NOT NULL default '0.00',
  icms_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cup,seq_ite)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcup20078'
#

CREATE TABLE tab_itcup20078 (
  seq_cup int(11) NOT NULL default '0',
  seq_ite char(8) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,2) NOT NULL default '0.00',
  icms_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cup,seq_ite)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcup2008'
#

CREATE TABLE tab_itcup2008 (
  seq_cup int(11) NOT NULL default '0',
  seq_ite char(8) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,2) NOT NULL default '0.00',
  icms_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cup,seq_ite)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcup2009'
#

CREATE TABLE tab_itcup2009 (
  seq_cup int(11) NOT NULL default '0',
  seq_ite char(8) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,2) NOT NULL default '0.00',
  icms_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cup,seq_ite)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcup2010'
#

CREATE TABLE tab_itcup2010 (
  seq_cup int(11) NOT NULL default '0',
  seq_ite char(8) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,2) NOT NULL default '0.00',
  icms_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cup,seq_ite)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcup2011'
#

CREATE TABLE tab_itcup2011 (
  seq_cup int(11) NOT NULL default '0',
  seq_ite char(8) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,2) NOT NULL default '0.00',
  icms_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_cup,seq_ite)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itcuptmp'
#

CREATE TABLE tab_itcuptmp (
  seq_cup int(11) NOT NULL default '0',
  seq_ite char(8) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  desconto_ite double(14,2) NOT NULL default '0.00',
  icms_ite double(14,2) NOT NULL default '0.00'
) TYPE=MyISAM;



#
# Table structure for table 'tab_itdev'
#

CREATE TABLE tab_itdev (
  NUM_DEV char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,2) NOT NULL default '0.00',
  SEQ_ITE int(11) NOT NULL default '0',
  PRIMARY KEY  (NUM_DEV,COD_PRO)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itdoc1'
#

CREATE TABLE tab_itdoc1 (
  seq_itd char(7) NOT NULL default '',
  seq_doc char(7) NOT NULL default '',
  Valor_itd double(14,2) default NULL,
  Reman_itd double(14,2) default NULL,
  Venc_itd date NOT NULL default '0000-00-00',
  juros_itd double(14,2) NOT NULL default '0.00',
  baixa_itd date default NULL,
  parc_itd int(11) NOT NULL default '0',
  desconto_itd double(14,2) NOT NULL default '0.00',
  cod_usu char(3) default NULL,
  PRIMARY KEY  (seq_itd)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itdocback'
#

CREATE TABLE tab_itdocback (
  seq_itd char(7) NOT NULL default '',
  seq_doc char(7) NOT NULL default '',
  Valor_itd double(14,2) default NULL,
  Reman_itd double(14,2) default NULL,
  Venc_itd date NOT NULL default '0000-00-00',
  juros_itd double(14,2) NOT NULL default '0.00',
  baixa_itd date default NULL,
  parc_itd int(11) NOT NULL default '0',
  desconto_itd double(14,2) NOT NULL default '0.00',
  cod_usu char(3) default NULL,
  PRIMARY KEY  (seq_itd)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itenspcp'
#

CREATE TABLE tab_itenspcp (
  seq_pcp char(6) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qfa_ipc double(14,2) default NULL,
  sts_ipc char(1) NOT NULL default '',
  upc_ipc double(14,2) default '0.00',
  suj_ipc double(14,2) NOT NULL default '0.00',
  total_ipc double(14,2) NOT NULL default '0.00',
  cod_for char(5) NOT NULL default '',
  PRIMARY KEY  (seq_pcp,cod_pro)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itmov'
#

CREATE TABLE tab_itmov (
  num_mov char(7) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  custo_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  seq_ite int(11) NOT NULL default '0',
  PRIMARY KEY  (num_mov,cod_pro)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itorc'
#

CREATE TABLE tab_itorc (
  num_orc char(7) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  unit_ite double(14,2) NOT NULL default '0.00',
  qtd_ite double(14,2) NOT NULL default '0.00',
  seq_ite int(11) NOT NULL default '0',
  PRIMARY KEY  (seq_ite,num_orc)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itped'
#

CREATE TABLE tab_itped (
  seq_ite int(11) NOT NULL default '0',
  seq_ped int(11) NOT NULL default '0',
  cod_pro char(8) NOT NULL default '',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,4) NOT NULL default '0.0000',
  cod_for char(5) NOT NULL default '',
  previsao_ite date NOT NULL default '0000-00-00',
  PRIMARY KEY  (seq_ite)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itperdas'
#

CREATE TABLE tab_itperdas (
  seq_per char(7) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  unit_itp double(14,2) default NULL,
  item_itp int(11) NOT NULL default '0',
  qtd_itp double(14,2) NOT NULL default '0.00'
) TYPE=MyISAM;



#
# Table structure for table 'tab_itsainf'
#

CREATE TABLE tab_itsainf (
  SEQ_SAI int(7) unsigned zerofill NOT NULL default '0000000',
  SEQ_ITS int(2) NOT NULL default '0',
  cod_pro char(8) NOT NULL default '',
  UNIT_ITS double(14,4) default NULL,
  QTD_ITS double(14,4) default NULL,
  icms_its double(14,2) NOT NULL default '0.00',
  cod_imp char(3) default NULL,
  PRIMARY KEY  (SEQ_SAI,SEQ_ITS)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itsainftmp'
#

CREATE TABLE tab_itsainftmp (
  SEQ_SAI int(7) unsigned zerofill NOT NULL default '0000000',
  SEQ_ITS int(2) NOT NULL default '0',
  cod_pro char(8) NOT NULL default '',
  UNIT_ITS double(14,4) default NULL,
  QTD_ITS double(14,4) default NULL,
  icms_its double(14,2) NOT NULL default '0.00',
  cod_imp char(3) default NULL
) TYPE=MyISAM;



#
# Table structure for table 'tab_ittrocas'
#

CREATE TABLE tab_ittrocas (
  seq_tro char(7) NOT NULL default '',
  cod_pro char(8) NOT NULL default '',
  qtdant_itt double(14,2) NOT NULL default '0.00',
  qtdatual_itt double(14,2) NOT NULL default '0.00',
  unit_itt double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_tro,cod_pro)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvendas'
#

CREATE TABLE tab_itvendas (
  NUM_VEN char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,3) NOT NULL default '0.000',
  SEQ_ITE int(11) NOT NULL default '0',
  descricao_ite char(50) NOT NULL default '',
  final_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (NUM_VEN,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvendas2007'
#

CREATE TABLE tab_itvendas2007 (
  NUM_VEN char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,3) NOT NULL default '0.000',
  SEQ_ITE int(11) NOT NULL default '0',
  descricao_ite char(50) NOT NULL default '',
  final_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (NUM_VEN,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvendas2008'
#

CREATE TABLE tab_itvendas2008 (
  NUM_VEN char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,3) NOT NULL default '0.000',
  SEQ_ITE int(11) NOT NULL default '0',
  descricao_ite char(50) NOT NULL default '',
  final_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (NUM_VEN,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvendas2009'
#

CREATE TABLE tab_itvendas2009 (
  NUM_VEN char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,3) NOT NULL default '0.000',
  SEQ_ITE int(11) NOT NULL default '0',
  descricao_ite char(50) NOT NULL default '',
  final_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (NUM_VEN,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvendas2010'
#

CREATE TABLE tab_itvendas2010 (
  NUM_VEN char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,3) NOT NULL default '0.000',
  SEQ_ITE int(11) NOT NULL default '0',
  descricao_ite char(50) NOT NULL default '',
  final_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (NUM_VEN,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvendas2011'
#

CREATE TABLE tab_itvendas2011 (
  NUM_VEN char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,3) NOT NULL default '0.000',
  SEQ_ITE int(11) NOT NULL default '0',
  descricao_ite char(50) NOT NULL default '',
  final_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (NUM_VEN,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvendas_copy'
#

CREATE TABLE tab_itvendas_copy (
  NUM_VEN char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,3) NOT NULL default '0.000',
  SEQ_ITE int(11) NOT NULL default '0',
  descricao_ite char(50) NOT NULL default '',
  final_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (NUM_VEN,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvendastmp'
#

CREATE TABLE tab_itvendastmp (
  NUM_VEN char(7) NOT NULL default '',
  COD_PRO char(8) NOT NULL default '',
  UNIT_ITE double(14,2) NOT NULL default '0.00',
  QTD_ITE double(14,3) NOT NULL default '0.000',
  SEQ_ITE int(11) NOT NULL default '0',
  descricao_ite char(50) NOT NULL default '',
  final_ite double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (NUM_VEN,SEQ_ITE)
) TYPE=MyISAM;



#
# Table structure for table 'tab_itvndcncl'
#

CREATE TABLE tab_itvndcncl (
  seq_itv char(7) NOT NULL default '',
  num_ven char(7) NOT NULL default '',
  seq_ite int(11) NOT NULL default '0',
  qtd_ite double(14,2) NOT NULL default '0.00',
  unit_ite double(14,2) NOT NULL default '0.00',
  cod_cai char(2) NOT NULL default '',
  cod_tur char(1) NOT NULL default '',
  data_ven datetime NOT NULL default '0000-00-00 00:00:00',
  cod_pro char(8) default NULL,
  PRIMARY KEY  (seq_itv)
) TYPE=MyISAM;



#
# Table structure for table 'tab_localche'
#

CREATE TABLE tab_localche (
  cod_loc char(2) NOT NULL default '',
  nome_loc char(20) NOT NULL default '',
  PRIMARY KEY  (cod_loc)
) TYPE=MyISAM;



#
# Table structure for table 'tab_marcas'
#

CREATE TABLE tab_marcas (
  cod_mar char(3) NOT NULL default '',
  nome_mar char(40) NOT NULL default '',
  PRIMARY KEY  (cod_mar)
) TYPE=MyISAM;



#
# Table structure for table 'tab_movint'
#

CREATE TABLE tab_movint (
  num_mov char(7) NOT NULL default '',
  tipo_mov char(1) NOT NULL default '',
  desc_mov char(40) NOT NULL default '',
  valor_mov double(14,2) NOT NULL default '0.00',
  data_mov date NOT NULL default '0000-00-00',
  situacao_mov char(1) NOT NULL default '',
  cod_usu char(3) NOT NULL default '',
  dataefet_mov date NOT NULL default '0000-00-00',
  PRIMARY KEY  (num_mov)
) TYPE=MyISAM;



#
# Table structure for table 'tab_notasai'
#

CREATE TABLE tab_notasai (
  seq_sai char(5) NOT NULL default '',
  nomedest_sai char(50) NOT NULL default '',
  cnpjdest_sai char(18) default NULL,
  enddest_sai char(30) default NULL,
  bairrodest_sai char(20) default NULL,
  cepdest_sai char(11) default NULL,
  munidest_sai char(40) default NULL,
  fonedest_sai char(16) NOT NULL default '',
  ufdest_sai char(2) NOT NULL default '',
  iedest_sai char(10) default NULL,
  emissao_sai date default NULL,
  datahora_sai datetime default NULL,
  BC_SAI double(14,2) NOT NULL default '0.00',
  VICMS_SAI double(14,2) NOT NULL default '0.00',
  PICMS_SAI double(14,2) NOT NULL default '0.00',
  bcsubs_SAI double(14,2) NOT NULL default '0.00',
  vsubs_SAI double(14,2) NOT NULL default '0.00',
  vipi_sai double(14,2) NOT NULL default '0.00',
  vfrete_sai double(14,2) NOT NULL default '0.00',
  vseg_sai double(14,2) NOT NULL default '0.00',
  totalprod_sai double(14,2) NOT NULL default '0.00',
  totalnota_sai double(14,2) NOT NULL default '0.00',
  nomefre_sai char(40) NOT NULL default '',
  contafre_sai char(1) default NULL,
  plveicfre_sai char(8) default NULL,
  ufveicfre_sai char(2) default NULL,
  cnpjfre_sai char(18) default NULL,
  endfre_sai char(50) default NULL,
  munifre_sai char(40) default NULL,
  uffre_sai char(2) default NULL,
  iefre_sai char(10) default NULL,
  qtdfre_sai char(10) default NULL,
  espfre_sai char(10) default NULL,
  numfre_sai char(10) default NULL,
  marcafre_sai char(10) default NULL,
  pbrutofre_sai char(10) default NULL,
  pliqfre_sai char(10) default NULL,
  dadosad_sai char(100) default NULL,
  PRIMARY KEY  (seq_sai)
) TYPE=MyISAM;



#
# Table structure for table 'tab_orc'
#

CREATE TABLE tab_orc (
  num_orc char(7) NOT NULL default '',
  cod_cli char(5) default NULL,
  nomecli_orc char(50) default NULL,
  valor_orc double(14,2) NOT NULL default '0.00',
  desconto_orc double(14,2) default NULL,
  data_orc datetime default '0000-00-00 00:00:00',
  validade_orc date default '0000-00-00',
  cod_usu char(3) default NULL,
  situacao_orc char(1) default NULL,
  obs_orc char(64) default NULL,
  PRIMARY KEY  (num_orc)
) TYPE=MyISAM;



#
# Table structure for table 'tab_parametros'
#

CREATE TABLE tab_parametros (
  NOME_EMP char(40) default NULL,
  ENDERECO_EMP char(40) default NULL,
  BAIRRO_EMP char(20) default NULL,
  CIDADE_EMP char(20) default NULL,
  UF_EMP char(2) default NULL,
  CEP_EMP char(10) default NULL,
  FONE_EMP char(16) default NULL,
  FAX_EMP char(14) default NULL,
  CNPJ_EMP char(18) default NULL,
  IE_EMP char(12) default NULL,
  NOMERESP_EMP char(40) default NULL,
  CPFRESP_EMP char(14) default NULL,
  RGRESP_EMP char(14) default NULL,
  NOMEMOEDA_PAR char(10) default NULL,
  SIMBMOEDA_PAR char(3) default NULL,
  PLURALMOEDA_PAR char(14) default NULL,
  TXFEDERATIVA_PAR double(14,2) default NULL,
  TXADM_PAR double(14,2) default NULL,
  TXDIVERSAS_PAR double(14,2) default NULL,
  MARGEMSEG_PAR double(14,2) default NULL,
  cod_loj char(3) NOT NULL default '001',
  msgecf_par char(48) default NULL,
  tipomsgecf_par char(1) default NULL,
  msgvenda_par char(200) default NULL,
  juros_par double(14,2) NOT NULL default '0.00',
  comissven_par double(14,2) NOT NULL default '0.00',
  fantasia_emp char(40) NOT NULL default '',
  IMPRESSORA_PAR char(1) NOT NULL default 'B',
  ipserver_par char(20) NOT NULL default '10.10.0.100',
  tipodoc_par char(1) NOT NULL default 'C',
  caminhoexec_par char(200) NOT NULL default '',
  plavendas_par char(3) NOT NULL default '',
  plarecdupl_par char(3) NOT NULL default '',
  plapagdupl_par char(3) NOT NULL default '',
  planaoclassc_par char(3) NOT NULL default '',
  planaoclassd_par char(3) NOT NULL default '',
  ipterminaltef_par char(20) default NULL,
  idestabtef_par char(20) default NULL
) TYPE=MyISAM;



#
# Table structure for table 'tab_pedidocompra'
#

CREATE TABLE tab_pedidocompra (
  seq_ped int(11) NOT NULL default '0',
  data_ped date NOT NULL default '0000-00-00',
  horaini_ped char(5) NOT NULL default '',
  horafec_ped char(5) NOT NULL default '',
  total_ped double(14,2) NOT NULL default '0.00',
  cod_usu char(3) NOT NULL default '',
  status_ped char(1) NOT NULL default '1',
  PRIMARY KEY  (seq_ped)
) TYPE=MyISAM;



#
# Table structure for table 'tab_perdas'
#

CREATE TABLE tab_perdas (
  seq_per char(7) NOT NULL default '',
  data_per date NOT NULL default '0000-00-00',
  cod_usu char(3) NOT NULL default '',
  situacao_per char(1) NOT NULL default '1',
  cod_for char(5) NOT NULL default '',
  PRIMARY KEY  (seq_per)
) TYPE=MyISAM;



#
# Table structure for table 'tab_planocontas'
#

CREATE TABLE tab_planocontas (
  cod_pla char(3) NOT NULL default '',
  nome_pla char(40) NOT NULL default '',
  tipo_Pla char(1) NOT NULL default '',
  PRIMARY KEY  (cod_pla)
) TYPE=MyISAM;



#
# Table structure for table 'tab_procomp'
#

CREATE TABLE tab_procomp (
  seq_pcp char(6) NOT NULL default '',
  dti_pcp date NOT NULL default '0000-00-00',
  dtf_pcp date NOT NULL default '0000-00-00',
  sts_pcp char(1) NOT NULL default '',
  pse_pcp int(11) NOT NULL default '0',
  pec_pcp int(11) NOT NULL default '0',
  ppc_pcp int(11) NOT NULL default '0',
  cod_cat char(2) default NULL,
  cod_mar char(3) default NULL,
  nome_pro char(60) default NULL,
  cod_usu char(3) NOT NULL default '',
  cod_for char(5) NOT NULL default '',
  PRIMARY KEY  (seq_pcp)
) TYPE=MyISAM;



#
# Table structure for table 'tab_prodpromo'
#

CREATE TABLE tab_prodpromo (
  cod_pro char(8) NOT NULL default '',
  datafin_pro date NOT NULL default '0000-00-00',
  PRIMARY KEY  (cod_pro)
) TYPE=MyISAM;



#
# Table structure for table 'tab_produtoatacado'
#

CREATE TABLE tab_produtoatacado (
  id int(11) NOT NULL auto_increment,
  codigo char(8) NOT NULL default '',
  qtd int(11) NOT NULL default '0',
  preco double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (id)
) TYPE=MyISAM;



#
# Table structure for table 'tab_produtos'
#

CREATE TABLE tab_produtos (
  cod_pro char(8) NOT NULL default '',
  codbarra_pro char(16) NOT NULL default '00000000',
  nome_pro char(50) default NULL,
  unidade_pro char(3) default NULL,
  cod_cat char(3) default NULL,
  cod_mar char(3) default NULL,
  codfab_pro char(30) default NULL,
  lucro_pro double(14,2) NOT NULL default '0.00',
  endereco_pro char(10) default NULL,
  er_pro double(14,2) NOT NULL default '0.00',
  ef_pro double(14,2) NOT NULL default '0.00',
  st_pro char(1) NOT NULL default 'T',
  cr_pro double(14,2) NOT NULL default '0.00',
  cf_pro double(14,2) NOT NULL default '0.00',
  cm_pro double(14,2) NOT NULL default '0.00',
  balanca_pro char(1) default 'N',
  pv_pro double(14,2) NOT NULL default '0.00',
  pp_pro double(14,2) NOT NULL default '0.00',
  cod_old char(10) NOT NULL default '0000000000',
  cod_imp char(3) NOT NULL default 'I05',
  mmd_pro double(14,4) NOT NULL default '0.0000',
  ipm_pro date default '0000-00-00',
  fpm_pro date default '0000-00-00',
  ce_pro double(14,2) NOT NULL default '0.00',
  ESTMIN_PRO double(14,2) NOT NULL default '0.00',
  ncm_pro char(10) default NULL,
  nfe_cst char(2) default '00',
  visivel_pro char(1) NOT NULL default 'S',
  PRIMARY KEY  (cod_pro),
  KEY nome (nome_pro),
  KEY codbarra (codbarra_pro)
) TYPE=MyISAM;



#
# Table structure for table 'tab_saidasnf'
#

CREATE TABLE tab_saidasnf (
  SEQ_SAI int(7) unsigned zerofill NOT NULL default '0000000',
  COD_CLI char(5) NOT NULL default '',
  NOTA_SAI int(12) unsigned zerofill NOT NULL default '000000000000',
  SERIE_SAI char(4) default NULL,
  NATOP_SAI char(10) default NULL,
  CFOP_SAI char(6) default NULL,
  VALOR_SAI double(14,4) NOT NULL default '0.0000',
  BC_SAI double(14,4) NOT NULL default '0.0000',
  VICMS_SAI double(14,4) NOT NULL default '0.0000',
  PICMS_SAI double(14,4) NOT NULL default '0.0000',
  VALP_SAI double(14,4) NOT NULL default '0.0000',
  COD_USU char(3) NOT NULL default '',
  EMISSAO_SAI date NOT NULL default '0000-00-00',
  situacao_sai char(1) NOT NULL default '1',
  despesas_Sai double(14,4) NOT NULL default '0.0000',
  num_ven char(7) NOT NULL default '',
  seq_cup int(10) unsigned NOT NULL default '0',
  nfe_sai int(7) NOT NULL default '0',
  frete_sai double(14,2) NOT NULL default '0.00',
  ipi_sai double(14,2) NOT NULL default '0.00',
  pis_sai double(14,2) NOT NULL default '0.00',
  cofins_sai double(14,2) NOT NULL default '0.00',
  seguro_sai double(14,2) NOT NULL default '0.00',
  desconto_sai double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (SEQ_SAI)
) TYPE=MyISAM;



#
# Table structure for table 'tab_saidastmp'
#

CREATE TABLE tab_saidastmp (
  SEQ_SAI int(7) unsigned zerofill NOT NULL default '0000000',
  COD_CLI char(5) NOT NULL default '',
  NOTA_SAI int(12) unsigned zerofill NOT NULL default '000000000000',
  SERIE_SAI char(4) default NULL,
  NATOP_SAI char(10) default NULL,
  CFOP_SAI char(6) default NULL,
  VALOR_SAI double(14,4) NOT NULL default '0.0000',
  BC_SAI double(14,4) NOT NULL default '0.0000',
  VICMS_SAI double(14,4) NOT NULL default '0.0000',
  PICMS_SAI double(14,4) NOT NULL default '0.0000',
  VALP_SAI double(14,4) NOT NULL default '0.0000',
  COD_USU char(3) NOT NULL default '',
  EMISSAO_SAI date NOT NULL default '0000-00-00',
  situacao_sai char(1) NOT NULL default '1',
  despesas_Sai double(14,4) NOT NULL default '0.0000',
  num_ven char(7) NOT NULL default '',
  seq_cup int(10) unsigned NOT NULL default '0',
  PRIMARY KEY  (SEQ_SAI)
) TYPE=MyISAM;



#
# Table structure for table 'tab_tipoven'
#

CREATE TABLE tab_tipoven (
  COD_TIP char(2) NOT NULL default '',
  PARCELAS_TIP int(11) NOT NULL default '0',
  JUROS_TIP double(14,4) NOT NULL default '0.0000',
  DESCRICAO_TIP char(40) NOT NULL default '',
  ENTRADA_TIP char(1) NOT NULL default '',
  INTERVALO_TIP int(11) NOT NULL default '0',
  tipo_tip char(1) NOT NULL default 'C',
  PRIMARY KEY  (COD_TIP)
) TYPE=MyISAM;



#
# Table structure for table 'tab_trocas'
#

CREATE TABLE tab_trocas (
  num_ven char(7) NOT NULL default '',
  data_tro datetime NOT NULL default '0000-00-00 00:00:00',
  cod_usu char(3) NOT NULL default '',
  seq_tro char(7) NOT NULL default '',
  abat_tro double(14,2) NOT NULL default '0.00',
  acres_tro double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (seq_tro)
) TYPE=MyISAM;



#
# Table structure for table 'tab_trtef'
#

CREATE TABLE tab_trtef (
  num_cpm char(6) NOT NULL default '',
  data_tra datetime NOT NULL default '0000-00-00 00:00:00',
  ip_cai char(20) NOT NULL default '',
  valor_tra double(14,2) NOT NULL default '0.00'
) TYPE=MyISAM;



#
# Table structure for table 'tab_turncai'
#

CREATE TABLE tab_turncai (
  cod_tur char(1) NOT NULL default '',
  cod_cai char(2) NOT NULL default '',
  data_tur date NOT NULL default '0000-00-00',
  login_usu char(3) NOT NULL default '',
  logout_usu char(3) NOT NULL default '',
  abertura_tur char(5) NOT NULL default '',
  fechamento_tur char(5) NOT NULL default '',
  valorabertura_tur double(14,2) NOT NULL default '0.00',
  valorfechamento_tur double(14,2) NOT NULL default '0.00',
  status_tur char(1) NOT NULL default 'A',
  fundo_tur double(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (cod_tur,cod_cai,data_tur)
) TYPE=MyISAM;



#
# Table structure for table 'tab_turnos'
#

CREATE TABLE tab_turnos (
  cod_tur char(1) NOT NULL default '',
  horaini_tur char(5) NOT NULL default '',
  horafin_tur char(5) NOT NULL default '',
  PRIMARY KEY  (cod_tur)
) TYPE=MyISAM;



#
# Table structure for table 'tab_usuarios'
#

CREATE TABLE tab_usuarios (
  COD_FUN char(3) NOT NULL default '',
  ULTACESSO_USU datetime default NULL,
  NIVEL_USU char(1) NOT NULL default '1',
  SENHA_USU char(10) default NULL,
  Status_USU char(1) NOT NULL default '1',
  PRIMARY KEY  (COD_FUN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vendas'
#

CREATE TABLE tab_vendas (
  NUM_VEN char(7) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DESCONTO_VEN double(14,2) default NULL,
  DATA_VEN datetime default '0000-00-00 00:00:00',
  COD_FUN char(3) NOT NULL default '',
  COD_USU char(3) NOT NULL default '',
  COD_TIP char(2) default NULL,
  SITUACAO_VEN char(1) NOT NULL default '',
  OBS_VEN char(64) default NULL,
  nomecli_ven char(50) NOT NULL default 'VENDA A CONSUMIDOR',
  cod_ven char(3) default '001',
  dataefet_ven date default '0000-00-00',
  COD_CAI char(2) NOT NULL default '01',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vendas2007'
#

CREATE TABLE tab_vendas2007 (
  NUM_VEN char(7) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DESCONTO_VEN double(14,2) default NULL,
  DATA_VEN datetime default '0000-00-00 00:00:00',
  COD_FUN char(3) NOT NULL default '',
  COD_USU char(3) NOT NULL default '',
  COD_TIP char(2) default NULL,
  SITUACAO_VEN char(1) NOT NULL default '',
  OBS_VEN char(64) default NULL,
  nomecli_ven char(50) NOT NULL default 'VENDA A CONSUMIDOR',
  cod_ven char(3) default '001',
  dataefet_ven date default '0000-00-00',
  COD_CAI char(2) NOT NULL default '01',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vendas2008'
#

CREATE TABLE tab_vendas2008 (
  NUM_VEN char(7) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DESCONTO_VEN double(14,2) default NULL,
  DATA_VEN datetime default NULL,
  COD_FUN char(3) NOT NULL default '',
  COD_USU char(3) NOT NULL default '',
  COD_TIP char(2) default NULL,
  SITUACAO_VEN char(1) NOT NULL default '',
  OBS_VEN char(64) default NULL,
  nomecli_ven char(50) NOT NULL default 'VENDA A CONSUMIDOR',
  cod_ven char(3) default NULL,
  dataefet_ven date default '0000-00-00',
  COD_CAI char(2) NOT NULL default '01',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vendas2009'
#

CREATE TABLE tab_vendas2009 (
  NUM_VEN char(7) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DESCONTO_VEN double(14,2) default NULL,
  DATA_VEN datetime default NULL,
  COD_FUN char(3) NOT NULL default '',
  COD_USU char(3) NOT NULL default '',
  COD_TIP char(2) default NULL,
  SITUACAO_VEN char(1) NOT NULL default '',
  OBS_VEN char(64) default NULL,
  nomecli_ven char(50) NOT NULL default 'VENDA A CONSUMIDOR',
  cod_ven char(3) default NULL,
  dataefet_ven date default '0000-00-00',
  COD_CAI char(2) NOT NULL default '01',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vendas2010'
#

CREATE TABLE tab_vendas2010 (
  NUM_VEN char(7) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DESCONTO_VEN double(14,2) default NULL,
  DATA_VEN datetime default NULL,
  COD_FUN char(3) NOT NULL default '',
  COD_USU char(3) NOT NULL default '',
  COD_TIP char(2) default NULL,
  SITUACAO_VEN char(1) NOT NULL default '',
  OBS_VEN char(64) default NULL,
  nomecli_ven char(50) NOT NULL default 'VENDA A CONSUMIDOR',
  cod_ven char(3) default NULL,
  dataefet_ven date default '0000-00-00',
  COD_CAI char(2) NOT NULL default '01',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vendas2011'
#

CREATE TABLE tab_vendas2011 (
  NUM_VEN char(7) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DESCONTO_VEN double(14,2) default NULL,
  DATA_VEN datetime default NULL,
  COD_FUN char(3) NOT NULL default '',
  COD_USU char(3) NOT NULL default '',
  COD_TIP char(2) default NULL,
  SITUACAO_VEN char(1) NOT NULL default '',
  OBS_VEN char(64) default NULL,
  nomecli_ven char(50) NOT NULL default 'VENDA A CONSUMIDOR',
  cod_ven char(3) default NULL,
  dataefet_ven date default '0000-00-00',
  COD_CAI char(2) NOT NULL default '01',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vendas_copy'
#

CREATE TABLE tab_vendas_copy (
  NUM_VEN char(7) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DESCONTO_VEN double(14,2) default NULL,
  DATA_VEN datetime default NULL,
  COD_FUN char(3) NOT NULL default '',
  COD_USU char(3) NOT NULL default '',
  COD_TIP char(2) default NULL,
  SITUACAO_VEN char(1) NOT NULL default '',
  OBS_VEN char(64) default NULL,
  nomecli_ven char(50) NOT NULL default 'VENDA A CONSUMIDOR',
  cod_ven char(3) default NULL,
  dataefet_ven date default NULL,
  COD_CAI char(2) NOT NULL default '01',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vendastmp'
#

CREATE TABLE tab_vendastmp (
  NUM_VEN char(7) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DESCONTO_VEN double(14,2) default NULL,
  DATA_VEN datetime default '0000-00-00 00:00:00',
  COD_FUN char(3) NOT NULL default '',
  COD_USU char(3) NOT NULL default '',
  COD_TIP char(2) default NULL,
  SITUACAO_VEN char(1) NOT NULL default '',
  OBS_VEN char(64) default NULL,
  nomecli_ven char(50) NOT NULL default 'VENDA A CONSUMIDOR',
  cod_ven char(3) default '001',
  dataefet_ven date default '0000-00-00',
  COD_CAI char(2) NOT NULL default '01',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_venfec'
#

CREATE TABLE tab_venfec (
  num_ven char(7) NOT NULL default '',
  avista_ven double(14,2) NOT NULL default '0.00',
  aprazo_ven double(14,2) NOT NULL default '0.00',
  finan_ven double(14,2) NOT NULL default '0.00',
  valorfinal_ven double(14,2) NOT NULL default '0.00',
  desconto_ven double(14,2) NOT NULL default '0.00',
  cod_usu char(3) NOT NULL default '',
  cod_tip char(2) default NULL,
  cod_fin char(2) default NULL,
  tipo_ven char(1) NOT NULL default '1',
  cod_cai char(2) default '01',
  cod_tur char(1) NOT NULL default '1',
  PRIMARY KEY  (num_ven)
) TYPE=MyISAM;



#
# Table structure for table 'tab_venfectmp'
#

CREATE TABLE tab_venfectmp (
  num_ven char(7) NOT NULL default '',
  avista_ven double(14,2) NOT NULL default '0.00',
  aprazo_ven double(14,2) NOT NULL default '0.00',
  finan_ven double(14,2) NOT NULL default '0.00',
  valorfinal_ven double(14,2) NOT NULL default '0.00',
  desconto_ven double(14,2) NOT NULL default '0.00',
  cod_usu char(3) NOT NULL default '',
  cod_tip char(2) default NULL,
  cod_fin char(2) default NULL,
  tipo_ven char(1) NOT NULL default '1',
  cod_cai char(2) default '01',
  cod_tur char(1) NOT NULL default '1',
  PRIMARY KEY  (num_ven)
) TYPE=MyISAM;



#
# Table structure for table 'tab_venfin'
#

CREATE TABLE tab_venfin (
  NUM_VEN char(7) NOT NULL default '',
  COD_FIN char(2) NOT NULL default '',
  VALOR_VEN double(14,2) NOT NULL default '0.00',
  DATA_VEN date NOT NULL default '0000-00-00',
  COD_USU char(3) NOT NULL default '',
  COD_CLI char(5) NOT NULL default '',
  AUTORIZACAO_VEN char(20) default NULL,
  baixa_ven date NOT NULL default '0000-00-00',
  status_ven char(1) NOT NULL default '1',
  PRIMARY KEY  (NUM_VEN)
) TYPE=MyISAM;



#
# Table structure for table 'tab_vinc'
#

CREATE TABLE tab_vinc (
  seq_vin char(7) NOT NULL default '',
  seq_itd char(7) NOT NULL default '',
  data_vin date NOT NULL default '0000-00-00',
  valor_vin double(14,2) NOT NULL default '0.00',
  acres_vin double(14,2) NOT NULL default '0.00',
  decres_vin double(14,2) NOT NULL default '0.00',
  venc_vin double(14,2) NOT NULL default '0.00',
  parc_vin double(14,2) NOT NULL default '0.00',
  seq_doc char(7) NOT NULL default '',
  PRIMARY KEY  (seq_vin)
) TYPE=MyISAM;

