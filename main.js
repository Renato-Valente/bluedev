console.log('Bom dia, flor do dia  😎')

import xlsx from 'xlsx';
import 'dotenv/config';
import pkg from 'pg';


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

    for(const row of rows){
        if(isNaN(row['Cod. Produto Final'])) continue;
        try{
            await pool.query(
                "insert into bd_produtos values($1, $2, $3, $4, $5, $6, $7, $8, $9, $10,$11)",
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
                "insert into bd_producao values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33)",
                row
            )
        }
        catch(err){
            console.log('erro:', err);
            continue;
        }

    }
}


//getProdutos();
getProducao();




//BD_Produtos
//BD_Poducao
