const { Pool } = require('pg');
const nodeExcel = require('excel-export');
const path = require('path');
const fs = require('fs');

const conf = {  
  stylesXmlFile : 'styles.xml',
  cols: [{
    caption: '字段',
    type: 'string',
    width: 28.7109375
  }, {
    caption: '类型',
    type: 'string',
    width: 28.7109375
  }, {
    caption: '长度',
    type: 'number',
    width: 28.7109375
  }, {
    caption: '最大长度',
    type: 'number',
    width: 28.7109375
  }, {
    caption: '允许为空',
    type: 'boolean',
    width: 28.7109375
  }, {
    caption: '注释',
    type: 'string',
    width: 48.7109375
  }],
};

module.exports = class Post2xlsx {

  constructor(options) {
    this.pool = new Pool({
      user: options.user ? options.user : 'postgres',
      host: options.host ? options.host : 'localhost',
      password: options.password ? options.password : 'postgres',
      port: options.port ? options.port : 5432,
      database: options.database,
    });
    this.output = options.output ? options.outout : path.resolve(process.cwd(), 'output');
    this.schema = options.schema ? options.schema : 'public';
  }

  async export() {
    await new Promise((resolve, reject) => {
      this.pool.query(`SELECT table_name FROM information_schema.tables WHERE table_schema = '${this.schema}'`, (err, res) => {
        if (err) {
          reject(err);
          return;
        }
        let rows = res.rows;
        rows.forEach((name) => {
          let sql = `
          SELECT
          A .attnum,
          A .attname AS field,
          T .typname AS TYPE,
          A .attlen AS LENGTH,
          A .atttypmod AS lengthvar,
          A .attnotnull AS NOTNULL,
          b.description AS COMMENT
        FROM
          pg_class C,
          pg_attribute A
        LEFT OUTER JOIN pg_description b ON A .attrelid = b.objoid
        AND A .attnum = b.objsubid,
        pg_type T
        WHERE
          C .relname = '${name.table_name}'
        AND A .attnum > 0
        AND A .attrelid = C .oid
        AND A .atttypid = T .oid
        ORDER BY
          A .attnum;
          `;
          conf.name = name.table_name;
          let count = rows.length;

          this.pool.query(sql, (err, res) => {
            if (err) {
              reject(err);
              return;
            }
            let rows = [];
            res.rows.forEach(row => {
              let list = [];
              list.push(row.field);
              list.push(row.type);
              list.push(row.length);
              list.push(row.lengthvar);
              list.push(row.notnull);
              list.push(row.comment);
              rows.push(list);
            });
            conf.rows = rows;
            let result = nodeExcel.execute(conf);
        
            if (!fs.existsSync(this.output)) {
              fs.mkdirSync(this.output);
            }
            fs.writeFileSync(path.resolve(this.output, name.table_name + '.xlsx'), result, 'binary');
            count = count - 1;
            console.log(count);
            if (count == 0) {
              resolve();
            }
          });
        }) 
      });
    });
  }

}