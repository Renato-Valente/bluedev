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
            delete(row['Id']);
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
                `insert into bd_parada (${cols.join(',')}) values($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
                Object.values(row)
            )
        }
        catch(err){
            console.log('deu errado:', err);
            continue;
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

    console.log('inserindo');
    const cols = [
        'id',
        'data',
        'sku',
        'descricao_prod',
        'pallets',
        'pacotes_cx',
        'total',
        'pedidos_aberto',
    ]
    let row = rows[0];
    await pool.query(
        `insert into bd_estoque values($1,$2,$3,$4,$5,$6,$7,$8)`,
        Object.values(row)
    )
}

const getFornecedor = async () => {
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

    //pegar pedidosAberto da outra aba

    console.log('vamos inserir 😎');
    const cols = [
        'nome_fornecedor',
        'codigo_fornecedor',
        'lead_time',
        'cond_pgto',
    ]


}


//getProdutos();
//getProducao();
getParadas();
//getEstoque();




//BD_Produtos
//BD_Poducao
