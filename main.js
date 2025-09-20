console.log('Bom dia, flor do dia  😎')

import xlsx from 'xlsx';
import 'dotenv/config';
import pkg from 'pg';



//UTILS
const MS_PER_DAY = 24 * 60 * 60 * 1000;

// converte número de dias desde 01/01/1900 para Date (UTC)
function dateFromDaysSince1900(n, { excel1900 = false } = {}) {
  if (typeof n !== 'number' || !isFinite(n)) {
    throw new TypeError('n precisa ser um número finito', n);
  }
  const days = Math.floor(n);            // ignoramos frações (horas)
  const epoch = Date.UTC(1900, 0, 1);    // 01/01/1900
  let offset = days - 1;                 // dia 1 => 1900-01-01

  // modo Excel (1900): ajusta o "dia fantasma" 29/02/1900
  if (excel1900 && days >= 60) offset -= 1;

  return new Date(epoch + offset * MS_PER_DAY);
}

// formata dd/mm/yy (dois dígitos)
function formatDDMMYY(date) {
  const dd = String(date.getUTCDate()).padStart(2, '0');
  const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
  const yy = String(date.getUTCFullYear());
  return `${yy}/${mm}/${dd}`;
}

// função principal
function days1900ToDDMMYY(n, { excel1900 = false } = {}) {
  // caso especial: Excel serial 60 deve mostrar 29/02/1900
  if (excel1900 && Math.floor(n) === 60) return '29/02/00';
  return formatDDMMYY(dateFromDaysSince1900(n, { excel1900 }));
}



const getProdutos = async () => {
    console.log('Vamos ler essas tabelas 🧐')
    const wb = xlsx.readFile('BaseDashProducao.xlsm', {
    // Ajuda quando houver células com datas
    cellDates: true
    });

    //console.log(wb.SheetNames);
    const aba = 'BD_Produtos';
    let rows = xlsx.utils.sheet_to_json(wb.Sheets[aba], {
    defval: null,          // mantém colunas com vazio = null
    raw: true              // mantém números como números
    });

    console.log('resultado:', rows[1]);

    //conectando com banco
    console.log('Vamos conectar com o banco 🐘');
    const { Pool } = pkg;
    
    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });


    //adicionar itens no banco
    console.log('Vamos inserir os dados na tabela 🙃');

    const cols = [
        'cod_produto_final',
        'nome_produto_final',
        'volumetria_ml',
        'volumetria_pct_cx',
        'qntd_por_pct_cx',
        'linha',
        'custo_prod',
        'custo_mp',
        'custo_emb',
        'embalagem_tipo',
        'ativo'
    ]

    for(const row of rows){
        if(isNaN(row['Cod. Produto Final'])) continue;
        try{
            await pool.query(
                "insert into bd_produtos values($1, $2, $3, $4, $5, $6, $7, $8, $9, $10,$11)\n" +
                "on conflict (cod_produto_final)\n"+
                `do update set(${cols.join(',')}) = ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10,$11)`,
                [
                    row['Cod. Produto Final'],
                    row['Nome Produto Final'],
                    row['Volumetria ML'],
                    row['Volumetria PCT/CX'],
                    row['Qntd por Pct/Cx'],
                    row['Linha'],
                    row['Custo Prod'],
                    row['Custo Mp'],
                    row['Custo Emb'],
                    row['Embalagem_Tipo'],
                    true
                ]
        )
        }
        catch(err){
            console.log(err);
            continue;
        }
    }

    console.log('acabamos por aqui. Kd o meu dinheiro? 🤑');
}

const getProducao = async () => {
    console.log('Vamos ler essas tabelas 🧐')
    const wb = xlsx.readFile('BaseDashProducao.xlsm', {
    // Ajuda quando houver células com datas
    cellDates: true
    });

    //console.log(wb.SheetNames);
    const aba = 'BD_Producao';
    let rows = xlsx.utils.sheet_to_json(wb.Sheets[aba], {
    defval: null,          // mantém colunas com vazio = null
    raw: true              // mantém números como números
    });

    console.log('resultado:', rows[1]);

    //conectando com banco
    console.log('Vamos conectar com o banco 🐘');
    const { Pool } = pkg;
    
    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });

    console.log('inserindo os dados ☺️');

    const cols = [
        'id',
        'data',
        'linha',
        'lote',
        'sku',
        'nome_produto',
        'volumetria_ml',
        'qntd_pct_cx',
        'producao_planejada',
        'producao_real',
        'hora_inicio_producao',
        'hora_fim_producao',
        'tempo_disp_producao',
        'tempo_inatividade',
        'total_producao',
        'perca_rot',
        'perca_tampa',
        'perca_garrafa_lata_vazia',
        'perca_garrafa_lata_cheia',
        'perca_preforma',
        'perca_filmeliso',
        'perca_filmestrech',
        'perca_filmeimpresso',
        'minimo',
        'turno',
        'lider',
        'lts_produzidos',
        'custo_prod',
        'custo_xarope',
        'custo_emb',
        'custo_prod_total',
        'custoemb_total',
        'customp_total',
    ]

    for(const r of rows){


        try{
            let row = Object.values(r);
            //Ajustando o row
            row[10] = row[10].toISOString().slice(11, 19); //converter para 13:23:04
            row[11] = row[11].toISOString().slice(11, 19);
            row[12] = row[12].toISOString().slice(11, 19);
            row[13] = row[13].toISOString().slice(11, 19);
            row[14] = 25;
            await pool.query(
                `insert into bd_producao (${cols.join(',')}) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33)\n`
                +`on conflict (id) do update set (${cols.join(',')}) = ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33)`,
                row
            )
        }
        catch(err){
            console.log('erro:', err);
            continue;
        }

    }
}

const getParadas = async () => {
    console.log('Vamos ler essas tabelas 🧐')
    const wb = xlsx.readFile('BaseDashParadas.xlsm', {
    // Ajuda quando houver células com datas
    cellDates: false
    });

    console.log(wb.SheetNames);
    const aba = 'BD_Parada';
    let rows = xlsx.utils.sheet_to_json(wb.Sheets[aba], {
    defval: null,          // mantém colunas com vazio = null
    raw: true              // mantém números como números
    });

    console.log('resultado:', rows[42]);

    //conectando com banco
    console.log('Vamos conectar com o banco 🐘');
    const { Pool } = pkg;
    
    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });


    console.log('vou inserir 🥒😨');

    const cols = [
        'id',
        'data',
        'lote',
        'sku',
        'hora_inicio',
        'hora_fim',
        'total_parado',
        'linha',
        'maquina',
        'motivo_parada',
        'categoria_parada',
        'custo_fabril',
        'custo_total',
    ]
    //rows = [rows[0]];
    console.log('lenght', rows.length)
    for(const r of rows /* let x = 0; x < 1; x++ */){

        try{


            let row = r;
            console.log('ESTE EH O OBJETO ANTES DAS ALTERACOES\n', row)
            //delete(row['Id']);
            delete(row['Nome do Produto']);
            //RENATOBOSS MEXENDO AQUI!!!
            //row['Data'] = '09/03/2000'
            row['Data'] = days1900ToDDMMYY(row['Data'], {excel1900: true})



            row['Total_Parado'] = Number(row['Total_Parado'].toFixed(7));
            row['Hora_Inicio'] = Number(row['Hora_Inicio'].toFixed(7));
            row['Hora_Fim'] = Number(row['Hora_Fim'].toFixed(7));
            /* row['Custo_Fabril'] = Number(row['Custo_Fabril'].toFixed(2));
            row['CustoTotal'] = Number(row['CustoTotal'].toFixed(2)); */
            //formatando row
            //row['Total_Parado'] = row['Total_Parado'].toISOString().slice(11, 19);
            console.log('size:', Object.values(row).length)
            console.log('ESSE EH O OBJETO QUE VAI SER INSERIDO NO BANCO\n:', row);


            await pool.query(
                `insert into bd_parada (${cols.join(',')}) values($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13)\n`
                +`on conflict (id, sku) do\n`
                +`update set (${cols.join(',')}) = ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13)`,
                Object.values(row)
            )
        }
        catch(err){
            console.log('deu errado:', err);
            continue;
        }
    }

}


const getFornecedor = async () => {
    console.log('Vamos ler essas tabelas 🧐')
    const wb = xlsx.readFile('01.08.2025 Custo dos Produtos.xlsm', {
    // Ajuda quando houver células com datas
    cellDates: true
    });

    console.log('sheets', wb.SheetNames)

    //console.log(wb.SheetNames);
    const aba = 'BD Fornecedor';
    let rows = xlsx.utils.sheet_to_json(wb.Sheets[aba], {
    defval: null,          // mantém colunas com vazio = null
    raw: false              // mantém números como números
    });

    //console.log('testando ', rows);

    //conectando com banco
    console.log('Vamos conectar com o banco 🐘');
    const { Pool } = pkg;
    
    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });


    console.log('vamos inserir 😎');
    
    for(const r of rows){
        try{
            let row = r; 
            const cols = [
                'nome_fornecedor',
                'codigo_fornecedor',
                'lead_time',
                'cond_pgto',
            ]

            console.log('ESTE EH O OBJETO ANTES DAS ALTERACOES\n', row);

            await pool.query(
                `insert into bd_fornecedor (${cols.join(',')}) values($1,$2,$3,$4)\n`
                +`on conflict (codigo_fornecedor) do\n`
                +`update set (${cols.join(',')}) = ($1,$2,$3,$4)`,
                Object.values(row)
            )
        }
        catch(err){
            console.log('deu tudo errado:', err);
        }
    }

}


const getEstoque = async() => {
    console.log('Vamos ler essas tabelas 🧐')
    const wb = xlsx.readFile('BaseEstoque.xlsm', {
    // Ajuda quando houver células com datas
    cellDates: true
    });

    //console.log(wb.SheetNames);
    const aba = 'BD_Estoque';
    let rows = xlsx.utils.sheet_to_json(wb.Sheets[aba], {
    defval: null,          // mantém colunas com vazio = null
    raw: true              // mantém números como números
    });

    //console.log('resultado:', rows[1]);

    //conectando com banco
    console.log('Vamos conectar com o banco 🐘');
    const { Pool } = pkg;
    
    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });


    //pegando o valor de PedidosAbertos

    console.log('Pegando o valor dos pedidos 😨');
    const pedidos = xlsx.utils.sheet_to_json(wb.Sheets['BD_PedidosAberto'], {
    defval: null,          // mantém colunas com vazio = null
    raw: true              // mantém números como números
    });

    console.log('teste:', pedidos[0])

    console.log('inserindo');
    const cols = [
        //'id',
        'data',
        'sku',
        'descricao_prod',
        'pallets',
        'pacotes_cx',
        'total',
        'pedidos_aberto',
    ]


    //inserindo o valor de pedidos
    for(const r of rows){
        try{
            let row = r;
            console.log('ESTE EH O OBJETO ANTES DAS ALTERACOES\n', row);

            const sku = row['SKU'];
            const pedido = pedidos.find((p) => p['SKU'] == sku);
            const abertos = pedido && pedido != null ? pedido['Total'] : 0;

            row['pedidos_abertos'] = abertos;

            console.log('DEPOIS DAS ALTERACOES\n', row);

            await pool.query(
                `insert into bd_estoque (${cols.join(',')}) values($1,$2,$3,$4,$5,$6,$7)`,
                Object.values(row)
            )
        }
        catch(err){
            console.log('deu tudo errado:', err);
        }
    }
}


const getEstrutura = async () => {


    console.log('Vamos ler essas tabelas 🧐')
    const wb = xlsx.readFile('01.08.2025 Custo dos Produtos.xlsm', {
    // Ajuda quando houver células com datas
    cellDates: true
    });

    console.log(wb.SheetNames);
    const aba = 'BD Preços dos Produtos';
    let rows = xlsx.utils.sheet_to_json(wb.Sheets[aba], {
    defval: null,          // mantém colunas com vazio = null
    raw: true              // mantém números como números
    });

    //conectando com banco
    console.log('Vamos conectar com o banco 🐘');
    const { Pool } = pkg;
    
    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });

    const cols = [
        'cod_produto_final',
        'codigo_materia_prima',
        'base_estrutura',
        'fator',
        'qnt_pacote_caixa',
    ]
    
    for(const r of rows){
        try{
            let row = r;
            await pool.query(
                //`insert into bd_estrutura (${cols.join(',')}) values($1,$2,$3,$4,$5)`,

                `insert into bd_estrutura (${cols.join(',')}) values($1,$2,$3,$4,$5)`
                +`on conflict (cod_produto_final, codigo_materia_prima) do\n`
                +`update set(${cols.join(',')}) = ($1,$2,$3,$4,$5)`,
                [
                    row['SKU Prod. Final'], row['Código Mat. Prima'], row['Base Estrutura'], row['Fator'], row['Quant Cons /pct']
                ]
            )
        }
        catch(err){
            console.log('deu errado aqui', err);
        }
    }
}

const getMateriaPrima = async () => {


    console.log('Vamos ler essas tabelas 🧐')
    const wb = xlsx.readFile('01.08.2025 Custo dos Produtos.xlsm', {
    // Ajuda quando houver células com datas
    cellDates: true
    });

    console.log(wb.SheetNames);
    const aba = 'BD Mat. Prima Original';
    let rows = xlsx.utils.sheet_to_json(wb.Sheets[aba], {
    defval: null,          // mantém colunas com vazio = null
    raw: true              // mantém números como números
    });

    //conectando com banco
    console.log('Vamos conectar com o banco 🐘');
    const { Pool } = pkg;
    
    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });

    const cols = [
        'sku',
        'data',
        'descricao_matprima',
        'um',
        'estoque'
    ]

    for(const r of rows){
        try{
            let row = r;
            console.log('ANTES DA MUDANÇA\n', row);

            await pool.query(
                `insert into bd_matprima (${cols.join(',')}) values($1,$2,$3,$4,$5)\n`
                +`on conflict (sku) do\n`
                +`update set (${cols.join(',')}) = ($1,$2,$3,$4,$5)`,
                [
                    row['Código Matéria Prima'], '09/03/2000',
                    row['Descrição'], row['UM'], row['Estoque']
                ]
            )
        }
        catch(err){
            console.log('deu errado 😢', err);
        }
    }
}

const getQualidade = async () => {
    console.log('Vamos ler essas tabelas 🧐')
    const wb = xlsx.readFile('BaseDashQualidade.xlsm', {
    // Ajuda quando houver células com datas
    cellDates: true
    });

    console.log(wb.SheetNames);
    const aba = 'BD_Qualidade';
    let rows = xlsx.utils.sheet_to_json(wb.Sheets[aba], {
    defval: null,          // mantém colunas com vazio = null
    raw: true              // mantém números como números
    });

    //conectando com banco
    console.log('Vamos conectar com o banco 🐘');
    const { Pool } = pkg;
    
    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });

    const cols = [
        'lote_id',
        'id_produto',
        'data_insercao',
        'data_analise',
        'conforme',
        'volume',
        'responsavel_analise',
        'agua_cloro',
        'agua_turbidez',
        'agua_ph',
        'xarope_brix',
        'xarope_acidez',
        'xarope_ph',
        'peso_garrafa',
        'milimetro_parede_garrafa',
        'analisevisual_garrafa',
        'co2_bebidafinal',
        'brix_bebidafinal',
        'sensorial_bebidafinal',
        'ph_bebidafinal',
        'acidez_bebidafinal'
    ]

    for(const r of rows){
        try{
            let row = r;
            console.log('ANTES DAS MUDANCAS\n', Object.values(row).length);
            console.log('ANTES DAS MUDANCAS\n', row);
            row['Conforme'] = row['Conforme'].toUpperCase().match('%SIM%') ? true : false;

            console.log('OBJETO DEPOIS DAS MUDANCAS\n', row);

            await pool.query(
                `insert into bd_qualidade (${cols.join(',')}) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21)\n`
                +`on conflict (lote_id) do \n`
                +`update set (${cols.join(',')}) = ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21)`,
                Object.values(row)
            )
        }
        catch(err){
            console.log('deu tudo errado', err);
        }
    }


}


//getProdutos(); //Done
//getProducao(); //Done
//getParadas(); //Done  coluna id adicionada e contraint (id, sku) criada
//getFornecedor(); // Done
//getEstrutura(); //Done
//getEstoque(); //(Done?) Sem id no Excel, usando o valor padrão (inscrement) sem update
//getMateriaPrima(); //(Done?) sem data no excel, usando 09/03/2000 em todos
//getQualidade(); 







//BD_Produtos
//BD_Poducao
