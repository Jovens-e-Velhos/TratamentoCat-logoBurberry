import pandas as pd
import openpyxl
import re

planilhaBrazil = pd.read_excel('Brazil Shipment A26 NO IL SHIPMENT 9 - TABELA DINAMICA.xlsx')
planilhaFabricantes = pd.read_excel('Operadores (2).xlsx')

def extracao_denominacao(self):
    if not isinstance(self, str):
        return self
    
    marca = re.search(rf'\bMARCA\b', self, flags=re.IGNORECASE)
    
    if not marca:
        return self.strip()
    
    pos_marca = marca.start()
    pos_separador = self.rfind(' - ', 0, pos_marca)
    
    if pos_separador == -1:
        return self[:pos_marca].strip()
    
    return self[:pos_separador].strip()

def codigo_fabricante(self, planilhaFabricantes):
    if not isinstance(self, str) is not self.strip():
        return None
    
    nome_busca = self.strip().upper()
    encontrados = planilhaFabricantes[planilhaFabricantes['Nome'].astype(str).str.strip().str.upper() == nome_busca]
    
    if encontrados.empty:
        return None
    
    return encontrados.iloc[0]['Código Interno']


def criacao_planilha_catalogo(planilhaOg):
    HEADERS_PRODUTOS = [
        'Código Interno', 'NCM', 'Denominação', 'Descrição', 'Modalidade',
        'Prioridade', 'Observação', 'Detalhamento Complementar', 'Ativo',
        'Mensagem Importação', 'REMOVER', 'Backlog', 'Novo Código Interno'
    ]

    HEADERS_IDENTIFICADORES = [
        'Produto(Código Interno)', 'Identificador', 'Mensagem Importação', 'REMOVER'
    ]

    HEADERS_FABRICANTES = [
        'Produto(Código Interno)', 'País', 'Operador Estrangeiro(Código Interno)',
        'CNPJ', 'Mensagem Importação', 'REMOVER'
    ]

    HEADERS_UNIDADES_NEGOCIO = [
        'Produto(Código Interno)', 'Unidade Negócio(Código Interno)',
        'Mensagem Importação', 'REMOVER'
    ]

    HEADERS_OPERADORES = [
        'Código Interno', 'Nome', 'Código TIN', 'País', 'Cidade',
        'Logradouro', 'Código Subdivisão País', 'CEP', 'E-mail',
        'Ativo', 'Mensagem Importação', 'Remover'
    ]
    
    HEADERS_OPERADORES_IDEN = [
        'Operador(Código Interno)', 'Número', 'Código', 'Mensagem Importação', 'Remover'
    ]
    
    produtos = []
    identificadores = []
    fabricantes = []
    unidade_negocio = []
    
    produtos_nao_cad = []
    identificadores_nao_cad = []
    fabricantes_nao_cad = []
    
    operadores_nao_cad = {}
    operadores_nao_cad_iden = {}
    
    for _, row in planilhaOg.iterrows():
        codigo_interno = row['ARTICLE #']
        ncm = row['Commodity Code'].replace('.','')
        detalhamento_complementar = row['Traducao']
        nome_fabricantes = row['Manufacturing name']
        
        
        produto_row = {
            'Código Interno'            : codigo_interno,
            'NCM'                       : ncm,
            'Denominação'               : extracao_denominacao(detalhamento_complementar),
            'Descrição'                 : detalhamento_complementar,
            'Modalidade'                : 'Importação',
            'Prioridade'                : 30,
            'Observação'                : None,
            'Detalhamento Complementar' : detalhamento_complementar,
            'Ativo'                     : 'Sim',
            'Mensagem Importação'       : None,
            'REMOVER'                   : None, 
            'Backlog'                   : None,
            'Novo Código Interno'       : None,
        }
        
        identificadores_row = {
            'Produto(Código Interno)'   : codigo_interno,
            'Identificador'             : codigo_interno,
            'Mensagem Importação'       : None,
            'REMOVER'                   : None, 
        }
        
        unidade_negocio.append({h: None for h in HEADERS_UNIDADES_NEGOCIO})
        
        codigo_fabricante_valor = codigo_fabricante(nome_fabricantes, planilhaFabricantes)
    
        if codigo_fabricante_valor is not None:
            produtos.append(produto_row)
            identificadores.append(identificadores_row)
            fabricantes.append({
                'Produto(Código Interno)'              : codigo_interno,
                'País'                                 : None,
                'Operador Estrangeiro(Código Interno)' : codigo_fabricante_valor,
                'CNPJ'                                 : None,
                'Mensagem Importação'                  : None,
                'REMOVER'                              : None,
            })
    
        else:
            produtos_nao_cad.append(produto_row)
            identificadores_nao_cad.append(identificadores_row)
            fabricantes_nao_cad.append({
                'Produto(Código Interno)'              : codigo_interno,
                'País'                                 : None,
                'Operador Estrangeiro(Código Interno)' : row['Manufacturer #'],
                'CNPJ'                                 : None,
                'Mensagem Importação'                  : None,
                'REMOVER'                              : None,
            })
            
            if nome_fabricantes not in operadores_nao_cad:
                operadores_nao_cad[nome_fabricantes] = {
                    'Código Interno'         : row['Manufacturer #'],
                    'Nome'                   : nome_fabricantes,
                    'Código TIN'             : None,
                    'País'                   : row['COO'],
                    'Cidade'                 : None,
                    'Logradouro'             : row['Manufacture address'],
                    'Código Subdivisão País' : None,
                    'CEP'                    : None,
                    'E-mail'                 : None,
                    'Ativo'                  : 'Sim',
                    'Mensagem Importação'    : None,
                    'Remover'                : None,
                }
                
                operadores_nao_cad_iden[nome_fabricantes] = {
                    'Operador(Código Interno)' : row['Manufacturer #'],
                    'Número'                   : row['Manufacturer #'],
                    'Código'                   : None,
                    'Mensagem Importação'      : None,
                    'Remover'                  : None,
                }
                
    df_produtos = pd.DataFrame(produtos, columns=HEADERS_PRODUTOS)
    df_identificadores = pd.DataFrame(identificadores, columns=HEADERS_IDENTIFICADORES)
    df_fabricantes = pd.DataFrame(fabricantes, columns=HEADERS_FABRICANTES)
    df_unidade_negocio = pd.DataFrame(unidade_negocio, columns=HEADERS_UNIDADES_NEGOCIO)
    
    df_produtos_naocad = pd.DataFrame(produtos_nao_cad, columns=HEADERS_PRODUTOS)
    df_identificadores_naocad = pd.DataFrame(identificadores_nao_cad, columns=HEADERS_IDENTIFICADORES)
    df_fabricantes_naocad = pd.DataFrame(fabricantes_nao_cad, columns=HEADERS_FABRICANTES)
    df_unidade_negocio_naocad = pd.DataFrame(columns=HEADERS_UNIDADES_NEGOCIO)
    
    df_operadores_naocad = pd.DataFrame(list(operadores_nao_cad.values()), columns=HEADERS_OPERADORES)
    df_identificadores_operadores_vazio = pd.DataFrame(list(operadores_nao_cad_iden.values()),columns=HEADERS_OPERADORES_IDEN)
    
    itens_nao_cadastrados = {
        'Produtos'                : df_produtos_naocad,
        'Identificadores'         : df_identificadores_naocad,
        'Fabricantes'             : df_fabricantes_naocad,
        'UnidadesNegocio'         : df_unidade_negocio_naocad,
        }
    
    fabricantes_nao_cadastrados = {
        'Operadores'       : df_operadores_naocad,
        'Identificadores'  : df_identificadores_operadores_vazio,
    }
    
    return (
        df_produtos, df_identificadores, df_fabricantes, df_unidade_negocio,
        itens_nao_cadastrados, fabricantes_nao_cadastrados
    )  

def salvar_planilhas(df_produtos, df_identificadores, df_fabricantes, df_unidades_negocio,
                      itens_nao_cadastrados, fabricantes_nao_cadastrados):
    
    with pd.ExcelWriter('ModeloProduto.xlsx', engine='openpyxl') as writer:
        df_produtos.to_excel(writer, sheet_name='Produtos', index=False)
        df_identificadores.to_excel(writer, sheet_name='Identificadores', index=False)
        df_fabricantes.to_excel(writer, sheet_name='Fabricantes', index=False)
        df_unidades_negocio.to_excel(writer, sheet_name='UnidadesNegocio', index=False)

    with pd.ExcelWriter('Itens não cadastrados.xlsx', engine='openpyxl') as writer:
        for nome_aba, df in itens_nao_cadastrados.items():
            df.to_excel(writer, sheet_name=nome_aba, index=False)

    with pd.ExcelWriter('Fabricantes não cadastrados.xlsx', engine='openpyxl') as writer:
        for nome_aba, df in fabricantes_nao_cadastrados.items():
            df.to_excel(writer, sheet_name=nome_aba, index=False)
    
    return


def remocao_campo(texto, campo):
    return re.sub(
        rf'\s+{campo}+:+\s*\S+',
        '',
        texto,
        flags=re.IGNORECASE
    )

def limpeza_tratamento (self):
    
    REGRAS = {
    'COMPRIMENTO': [
        'CASACO', 'CARDIGA', 'CAMISETA', 'CAMISA',
        'BOLSA', 'PORTA CARTAO', 'CALCA', 'SUETER',
        'ECHARPE', 'LENCO', 'CINTO',
    ],
    'FORRO': [
        'ECHARPE', 'LENCO', 'CINTO',
        'OCULOS', 'VELA',
    ],
    'MANGA': ['OCULOS', 'VELA'],
    'MALHA': ['OCULOS', 'VELA', 'BRACELETE', 'BRINCO', 'COLAR'],
    'SOLA': [
        'CASACO', 'CARDIGA', 'CAMISETA', 'CAMISA',
        'BOLSA', 'PORTA CARTAO', 'CALCA', 'SUETER',
        'ECHARPE', 'LENCO', 'CINTO',
        'OCULOS', 'VELA',
    ],
}
    
    PADRAO = {
        '(CANADA ONLY)', '(FOR FILLING ONLY)',
    }
    
    texto = self
    texto_upper = texto.upper()
    padrao = '|'.join(map(re.escape, PADRAO))
    try:
        
        for campos, categoria in REGRAS.items():
            if any(cat in texto_upper for cat in categoria):
                texto = remocao_campo(texto, campos)
        tratamento = re.sub(r'\s+\w+:+\s*nan\b','',texto)
        tratamento = re.sub(r'(\s*-\s*)+', ' - ', tratamento)
        tratamento = re.sub(padrao, '',tratamento, flags=re.IGNORECASE)
        tratamento = tratamento.strip(' -').strip() 
              
        return tratamento
    except Exception as e:
        print(f'Erro encontrado: {e}')

def traducao (self):
    
    DICIONARIO_TRADUCAO = {
    'O-MINI BAG'                    : 'BOLSA',
    'SQUARE SCARF'                  : 'LENÇO QUADRADO',
    '1 PIECE SWIMMING SUIT'         : 'MAIO (MALHA OU TECIDO)',
    'ACCESSOIRE VOYAGE'             : 'ACESSORIO DE VIAGEM',
    'ACETATE'                       : 'ACETATO',
    'ACRYLIC'                       : 'ACRILICO',
    'ALUMINUM'                      : 'ALUMINIO',
    'AUTRE FORME SANS CHAINE'       : 'POCHETE',
    'BACK PACK'                     : 'MOCHILA',
    'BAG BOX & EVENING PIECES'      : 'BOLSA',
    'BARRETTE'                      : 'PRESILHA DE CABELO',
    'BASE DU DESSUS'                : 'BASE SUPERIOR',
    'BATH TOWEL'                    : 'TOALHA DE BANHO',
    'BLOUSE'                        : 'BLUSA OU CAMISA (MALHA OU TECIDO)',
    'BLOUSON'                       : 'JAQUETA',
    'BOATER'                        : 'CHAPEU DE BARQUEIRO',
    'BOTTES'                        : 'BOTAS',
    'BOUCLE'                        : 'FIVELA',
    'BOUTONNAGE'                    : 'ABOTOAMENTO',
    'BOWLING BAG'                   : 'BOLSA',
    'BRACELET'                      : 'BRACELETE',
    'BRAID'                         : 'TRANÇA',
    'BRASS'                         : 'LATAO',
    'BRODERIE'                      : 'BORDADO',
    'BROOCH'                        : 'BROCHE',
    'BUCKET BAG'                    : 'BOLSA',
    'BUTTERFLY SUNGLESSES'          : 'OCULOS DE SOL',
    'BUTTERFLY SUNVIDROES'          : 'OCULOS DE SOL',
    "CALF'S SKIN"                   : 'COURO DE BEZERRO',
    'CALFSKIN'                      : 'COURO DE BEZERRO',
    'CAMELLIA BROCHE'               : 'BROCHE',
    'CAPE'                          : 'CAPA',
    'CARROT PANT'                   : 'CALÇA',
    'CASHMERE'                      : 'CAXEMIRA',
    'CAT EYE SUNGLASSES'            : 'OCULOS DE SOL',
    'CHOUCHOU'                      : 'ELASTICO DE CABELO',
    'CIGARETTE PANT'                : 'CALÇA',
    'CLOCHE'                        : 'CHAPEU',
    'CLUTCH BAG'                    : 'BOLSA DE MAO',
    'COAT'                          : 'CASACO COMPRIDO',
    'COL+POIGNETS+BASDECORPS'       : 'GOLA + PUNHOS + PARTE INFERIOR DO CORPO',
    'COLLAR'                        : 'COLARINHO',
    'COMMON METAL'                  : 'METAL COMUM',
    'COTTON'                        : 'ALGODAO',
    'CROUTE DE VEAU'                : 'CRISTA DE VITELA',
    'CUFFS'                         : 'PUNHOS',
    'CUSHION'                       : 'ALMOFADA',
    'DIVERS MERCH'                  : 'IMÃ DECORATIVO',
    'DOUBLURE ELEMENT'              : 'REVESTIMENTO DE ELEMENTOS',
    'DOUBLURE POCHE'                : 'FORRO DO BOLSO',
    'DRESS'                         : 'VESTIDO',
    'DYED LAMB SHEARLING'           : 'COURO DE CORDEIRO TINGIDO',
    'DYED SHEEP SHEARLING'          : 'COURO DE OVELHA TINGIDA',
    'EARRINGS'                      : 'BRINCOS',
    'ELASTANE'                      : 'ELASTANO',
    'ELEMENT 1'                     : 'ELEMENTO 1',
    'ELEMENT 2'                     : 'ELEMENTO 2',
    'ELEMENT 3'                     : 'ELEMENTO 3',
    'ELEMENT 4'                     : 'ELEMENTO 4',
    'ELEMENT 5'                     : 'ELEMENTO 5',
    'EMBROIDERY SUPPORT'            : 'SUPORTE DO BORDADO',
    'EMBROIDERY'                    : 'BORDADO',
    'EMMANCHURE'                    : 'DETALHE',
    'ESPADRILLES'                   : 'ALPARGATAS',
    'ETHYLENE-VINYL ACETATE'        : 'ACETATO DE ETILENO-VINILA',
    'ETHYLENE-VINYL ACETATO'        : 'ETILENO-VINIL ACETATO',
    'FANCY BELT'                    : 'CINTO DE METAL',
    'FASH BTQ HANGER'               : 'CABIDE',
    'FAUX FUR'                      : 'COURO SINTETICO',
    'FEATHERS'                      : 'PENAS',
    'FIL DE COUTURE'                : 'LINHA DA COSTURA',
    'FINISHING'                     : 'ACABAMENTO',
    'FLAPBAG WITH HANDLE'           : 'BOLSA COM ALÇA',
    'FLAPBAG'                       : 'BOLSA',
    'FLATS'                         : 'SAPATILHAS',
    'FRAME'                         : 'ARMAÇÃO',
    'FRESH WATER PEARL'             : 'PÉROLA DE ÁGUA DOCE',
    'FULL MANNEQUIN'                : 'MANEQUIM',
    'GARNET'                        : 'GRANADA',
    'GILET'                         : 'COLETE',
    'GLASS'                         : 'VIDRO',
    'GLOVES'                        : 'LUVAS (MALHA OU TECIDO)',
    'GOATSKIN'                      : 'COURO DE CABRA',
    'GUM'                           : 'GOMA',
    'HAIRCLIP'                      : 'PRESILHA DE CABELO',
    'HANDBAGS'                      : 'BOLSA DE MAO',
    'HANGER'                        : 'CABIDE',
    'HOBO BAG'                      : 'BOLSA',
    'HOLDER'                        : 'SUPORTE',
    'INSET'                         : 'INSERIDO',
    'INSOLE'                        : 'PALMILHAS',
    'JACKET'                        : 'CASACO OU BLAZER',
    'JEANS'                         : 'CALÇA',
    'JUMP SUIT'                     : 'MACACAO (MALHA OU TECIDO)',
    'JUTE'                          : 'JUTA',
    'KNIT TOP'                      : 'TOP DE MALHA',
    'LACE UP'                       : 'BOTAS',
    'LAMBSKIN'                      : 'COURO DE CORDEIRO',
    'LARGE SHOPPING >30CM'          : 'BOLSA',
    'LEATHER WOMAN BELT'            : 'CINTO',
    'LINEN'                         : 'LINHO',
    'LINING'                        : 'FORRO',
    'L-LONG WALLET'                 : 'CARTEIRA',
    'MAGNETIZED MATERIAL'           : 'MATERIAL MAGNETIZADO',
    'MANNQ ACCESS'                  : 'MANEQUIM',
    'MANNQ PARTS'                   : 'MANEQUIM',
    'MANNQ SET PARTS'               : 'MANEQUIM',
    'MARY JANES'                    : 'SAPATILHAS',
    'MED. DENSITY FIBERBOARD'       : 'MDF',
    'MESSENGER BAG'                 : 'BOLSA MENSAGEIRO',
    'METALLISED FIBER'              : 'FIBRA METALIZADA',
    'MISCELLANEOUS BAG'             : 'BOLSA',
    'M-MEDIUM WALLET'               : 'CARTEIRA',
    'MULES'                         : 'MULES',
    'NATURAL FIBER'                 : 'FIBRA NATURAL',
    'NECKLACE'                      : 'COLAR',
    'O-ACCESSOIRE DE SAC'           : 'BOLSA',
    'O-CARD HOLDER'                 : 'PORTA-CARTÃO',
    'O-COIN PURSE'                  : 'PORTA-MOEDAS',
    'O-OTHER BELT BAG'              : 'BOLSA CINTO NCM 4202.9',
    'O-OTHER SHAPE WITH CHAIN'      : 'BOLSA',
    'O-PORTE MONNAIE A BAND'        : 'BOLSA',
    'O-PURSEVANITY'                 : 'BOLSA',
    'OTHER FIBRES'                  : 'OUTRAS FIBRAS',
    'OTHER OPEN SHOES'              : 'SAPATO ABERTO',
    'OUTER SOLE'                    : 'SOLA',
    'OVAL SUNGLASSES'               : 'OCULOS DE SOL',
    'O-VANITY A BANDOULIERE'        : 'BOLSA',
    'PADDING'                       : 'PREENCHIMENTO',
    'PANTOS SUNVIDROES'             : 'OCULOS DE SOL',
    'PANTS'                         : 'CALÇA',
    'PAPER'                         : 'PAPEL',
    'PAREMENTURE'                   : 'FRENTE',
    'PARKA'                         : 'PARKA',
    'PATTE DE POCHE'                : 'ABA DE BOLSO',
    'PATTE'                         : 'CARCETE',
    'PEARL'                         : 'PEROLA',
    'PEROLAY PEROLAY'               : 'PEROLA PEROLA',
    'PHEASANT FEATHER'              : 'PENA DE FAISÃO',
    'PHONE HOLDER WITH CHAIN'       : 'PORTA CELULAR',
    'PILOT SUNVIDROES'              : 'OCULOS DE SOL',
    'PLASTIC'                       : 'PLASTICO',
    'PLATEAU'                       : 'PLATO',
    'POCKETING FABRIC'              : 'TECIDO DO BOLSO',
    'POLYAMIDE'                     : 'POLIAMIDA',
    'POLYESTER'                     : 'POLIESTER',
    'POLYETHYLENE TEREPHTALATE'     : 'TEREFTALATO DE POLIETILENO',
    'POLYMETHYL METHACRYLATE'       : 'METACRILATO DE POLIMETILA',
    'POLYPROPYLENE'                 : 'POLIPROPILENO',
    'POLYURETHANE THERMOPLAST.'     : 'POLIURETANO TERMOPLASTICO',
    'POLYURETHANE'                  : 'POLIURETANO',
    'PORTE MONNAIE A CHAINE'        : 'BOLSA',
    'PORTE MONNAIE SANS CHAINE'     : 'CARTEIRA',
    'REAL FUR'                      : 'COURO AUTENTICO',
    'RECTANGLE SUNVIDROES'          : 'OCULOS DE SOL',
    'RESIN'                         : 'RESINA',
    'RHINESTONE'                    : 'STRASS',
    'RIBBINGS'                      : 'NERVURAS',
    'RING'                          : 'ANEL',
    'RISER'                         : 'DISPLAY EXPOSITOR',
    'ROCK CRYSTAL'                  : 'PEDRA DE CRISTAL',
    'ROUND SUNVIDROES'              : 'OCULOS DE SOL',
    'RUBBER'                        : 'BORRACHA',
    'SANDALS'                       : 'SANDALIAS',
    'SCARF'                         : 'CACHECOL',
    'SET DE PLAGE'                  : 'BOLSA DE MAO',
    'SHEEPSKIN'                     : 'COURO DE CARNEIRO',
    'SHIELD SUNVIDROES'             : 'OCULOS DE SOL',
    'SHIRT'                         : 'CAMISA',
    'SHOES'                         : 'SAPATO',
    'SHORT BOOTS'                   : 'BOTAS',
    'SHORT FORRO'                   : 'FORRO CURTO',
    'SILK'                          : 'SEDA',
    'SKIRT'                         : 'SAIA',
    'SLEEVES'                       : 'MANGAS',
    'SLG ON CHAIN'                  : 'BOLSA',
    'SLIM BANDEAU'                  : 'BANDEAU PEQUENO',
    'SLINGS'                        : 'SANDALIAS',
    'SMALL LEATHER GOODS'           : 'BOLSA',
    'SMALL SHOPPING <31CM'          : 'BOLSA',
    'SMALL WALLET'                  : 'CARTEIRA PEQUENA',
    'SNEAKERS'                      : 'TENIS',
    'SQUARE CACHECOL'               : 'CACHECOL',
    'SQUARE SUNVIDROES'             : 'OCULOS DE SOL',
    'S-SMALL CASE'                  : 'BOLSA PEQUENA',
    'STAINLESS STEEL'               : 'AÇO INOXIDÁVEL',
    'STOLE'                         : 'ESTOLA',
    'STRAIGHT PANT'                 : 'CALÇA RETA',
    'STRAP'                         : 'ALÇA',
    'STRAW'                         : 'PALHA',
    'SUNGLASSES'                    : 'OCULOS DE SOL',
    'SUPPORT'                       : 'SUPORTE',
    'SWIMMING TOP'                  : 'TOP DE NATAÇÃO',
    'SWIMMING TRUNK'                : 'SHORT DE BANHO',
    'TEE-SHIRT'                     : 'CAMISA',
    'THERMOPLAST.'                  : 'TERMOPLASTICO',
    'THONGS'                        : 'SANDALIAS',
    'TOP'                           : 'TOP',
    'TRAY'                          : 'BANDEJA',
    'TRUNK FORRO'                   : 'FORRO',
    'UPPER'                         : 'SUPERIOR',
    'UPSIDE'                        : 'SUPERIOR',
    'VANITY A CHAINE'               : 'BOLSA',
    'VEAU VELOURS'                  : 'CAMURÇA DE BEZERRO',
    'VELOURS SYNTHETIQUE'           : 'VELUDO',
    'VERRES'                        : 'LENTES',
    'VEST'                          : 'COLETE',
    'WATER PAINT'                   : "TINTA A BASE D'AGUA",
    'WIDE LEG PANT'                 : 'CALÇA LARGA',
    'WITH HANDLE'                   : 'COM ALÇA',
    'WOOD'                          : 'MADEIRA',
    'WOOL'                          : 'LA',
    'Long'                          : 'LONGO',
    'WOVEN'                         : 'TECIDO PLANO',
    'Straight'                      : 'RETÍLINEA',
    'FORRO: Full'                   : 'FORRO COMPLETO',
    'CALF GRAIN LEATHER (BOS TAURUS)':'COURO BOVINO',
    'COW GRAIN LEATHER (BOS TAURUS)': 'COURO BOVINO',
    'Above knee'                    : 'ACIMA DO JOELHO',
    'Half lined'                    : 'MEIO FORRO',
    'LAMB LEATHER (OVIS ARIES ARIES)':'COURO DE CORDEIRO'
}
    termos_ordenados = sorted(DICIONARIO_TRADUCAO.keys(), key=len, reverse=True)
    resposta = self
    
    for termo in termos_ordenados:
        
        if re.search(rf'\b{re.escape(termo)}\b', self, flags=re.IGNORECASE):
            print(f'BATEU: {termo} > {DICIONARIO_TRADUCAO[termo]}')
        
        resultado = DICIONARIO_TRADUCAO[termo]
        resposta = re.sub(
        rf'\b{re.escape(termo)}\b',
        resultado,
        resposta,
        flags=re.IGNORECASE
    )
        
    return resposta

def principal():
    
    planilhaBrazil.drop_duplicates(subset=['ARTICLE #'], inplace=True)   
    planilhaBrazil['RES'] = planilhaBrazil.apply(lambda i: f"SKU: {i['ARTICLE #']}- Cód:{i['PrdCode']} - {i['Descrição português']} - MARCA BURBERRY - {i['Composition']} - FORRO: {i['Lining Full or Partial']} COMPRIMENTO: {i['Above, below on knee']} MANGA: {i['Sleeve Type']} - MALHA: {i['Knit Type']} - SOLA: {i['Sole Composition']}", axis=1)
    planilhaBrazil['Tratamento'] = planilhaBrazil['RES'].apply(limpeza_tratamento)
    planilhaBrazil['Traducao'] = planilhaBrazil['Tratamento'].apply(traducao)
    resposta = criacao_planilha_catalogo(planilhaBrazil)
    salvar_planilhas(*resposta)
    print(resposta)
    #resposta = planilhaBrazil.to_excel('tratado.xlsx', index=False)
        
if __name__ == '__main__':
    principal()