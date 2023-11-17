// @ts-nocheck
const router = require('express-promise-router').default();
const graph = require('../graph.js');
const { body } = require('express-validator');

const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const fs = require('fs');
const path = require('path');

router.get('/',
  async function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/');
    } else {
      let params = {};
      try {
        const file = await graph.getFiles(
          req.app.locals.msalClient,
          req.session.userId);

        params = file;
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not fetch file',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
      }

      // console.log(params);
      res.render('files', {files: params});
    }
  }
);

router.get('/generate',
  function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/');
    } else {
      res.render('generate');
    }
  }
);

router.post('/generate',[
  body('nama_kota').escape(),
  body('tanggal_pembuatan_dokumen').escape(),
  body('judul_dokumen').escape(),
  body('nomor_dokumen').escape(),
  body('kode_departemen').escape(),
  body('bulan_ini').escape(),
  body('tahun_ini').escape(),
  body('tanggal_mulai_dokumen').escape(),
  body('tanggal_selesai_dokumen').escape(),
  body('isi_dokumen').escape(),
  body('nama_pihak_pertama').escape(),
  body('nama_pihak_kedua').escape()
], async function(req, res) {
  if (!req.session.userId) {
    res.redirect('/');
  } else {
    const content = fs.readFileSync(
      path.resolve('./template', 'Template.docx'),
      'binary'
    );
      
    const zip = new PizZip(content);
      
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });        

    doc.render({
      nama_kota: req.body.nama_kota,
      tanggal_pembuatan_dokumen: req.body.tanggal_pembuatan_dokumen,
      judul_dokumen: req.body.judul_dokumen,
      nomor_dokumen: req.body.nomor_dokumen,
      kode_departemen: req.body.kode_departemen,
      bulan_ini: req.body.bulan_ini,
      tahun_ini: req.body.tahun_ini,
      tanggal_mulai_dokumen: req.body.tanggal_mulai_dokumen,
      tanggal_selesai_dokumen: req.body.tanggal_selesai_dokumen,
      isi_dokumen: req.body.isi_dokumen,
      nama_pihak_pertama: req.body.nama_pihak_pertama,
      nama_pihak_kedua: req.body.nama_pihak_kedua
    });

    var buf = doc.getZip().generate({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    });
      
    fs.writeFileSync(path.resolve('./template', req.body.nomor_dokumen+'.docx'), buf);

    try {
      const response = await graph.upload(
        req.app.locals.msalClient,
        req.session.userId,
        req.body.nomor_dokumen);

      console.log(response);
      // res.redirect('/onedrive');
      res.json({ message : response.id , link : response.webUrl});
    } catch (err) {
      req.flash('error_msg', {
        message: 'Could not fetch file',
        debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
      });
    }
  }
}
);

module.exports = router;
