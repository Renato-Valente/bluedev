console.log('Vamos conectar com o banco 🐘');
import 'dotenv/config';
import pkg from 'pg';

(async() => {
    // db.js
    const { Pool } = pkg;

    const pool = new Pool({
    host: process.env.PGHOST,
    port: Number(process.env.PGPORT || 5432),
    database: process.env.PGDATABASE,
    user: process.env.PGUSER,
    password: process.env.PGPASSWORD,
    //  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false, // ajuste em prod
    });

    const teste = await pool.query(
        'select * from bd_produtos'
        //"insert into bd_produtos values(3,'testinho',12000,12000,10,'linha teste',10.00,10.00,10.00,'embas',true);"
    )

    console.log('teste:', teste.rows);
    process.exit(0);
})()

