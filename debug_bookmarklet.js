function bookmarklet() {
    var s = document.createElement('script');
    s.src = 'https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js';
    s.onload = function () {
        var css = '#gi-overlay{position:fixed;top:20px;right:20px;width:320px;background:white;border:1px solid #ccc;box-shadow:0 4px 12px rgba(0,0,0,0.15);z-index:10000;padding:20px;font-family:sans-serif;border-radius:8px;font-size:14px}#gi-header{font-weight:bold;font-size:16px;margin-bottom:15px;display:flex;justify-content:space-between;border-bottom:1px solid #eee;padding-bottom:10px}.gi-label{display:block;margin-bottom:4px;font-weight:500;color:#555}.gi-select{width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;margin-bottom:10px}.gi-btn{background:#2563eb;color:white;border:none;padding:10px;border-radius:4px;cursor:pointer;width:100%;font-weight:bold;margin-top:10px}.gi-btn:disabled{background:#ccc;cursor:not-allowed}#gi-status{margin-top:15px;font-size:13px;color:#666;line-height:1.4;background:#f9fafb;padding:10px;border-radius:4px}';
        var style = document.createElement('style');
        style.textContent = css;
        document.head.appendChild(style);
        var d = document.createElement('div');
        d.id = 'gi-overlay';
        // Potential syntax error source: nested quotes in innerHTML
        d.innerHTML = '<div id=\'gi-header\'><span>Import Grades</span><span onclick=\'this.parentElement.parentElement.remove()\' style=\'cursor:pointer\'>&times;</span></div><input type=\'file\' id=\'gi-file\' accept=\'.xlsx\' style=\'width:100%\'/><div id=\'gi-map\' style=\'display:none\'><div class=\'gi-row\'><label class=\'gi-label\'>Student Column</label><select id=\'gi-col-name\' class=\'gi-select\'></select></div><div class=\'gi-row\'><label class=\'gi-label\'>Grade Column</label><select id=\'gi-col-grade\' class=\'gi-select\'></select></div></div><button id=\'gi-btn\' class=\'gi-btn\' disabled>Import Grades</button><div id=\'gi-msg\'>Select an Excel file to start.</div>';
        document.body.appendChild(d);
        var wb, json;
        document.getElementById('gi-file').onchange = function (e) {
            var f = e.target.files[0];
            if (!f) return;
            var r = new FileReader();
            r.onload = function (e) {
                wb = XLSX.read(e.target.result, {
                    type: 'array'
                });
                json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
                if (json.length) {
                    var h = Object.keys(json[0]);
                    var sn = document.getElementById('gi-col-name');
                    var sg = document.getElementById('gi-col-grade');
                    [sn, sg].forEach(s => {
                        s.innerHTML = '';
                        h.forEach(o => {
                            var opt = document.createElement('option');
                            opt.value = o;
                            opt.text = o;
                            s.add(opt)
                        })
                    });
                    document.getElementById('gi-map').style.display = 'block';
                    document.getElementById('gi-btn').disabled = false;
                    document.getElementById('gi-msg').textContent = 'Columns mapped. Ready to import.';
                } else {
                    document.getElementById('gi-msg').textContent = 'Empty file';
                }
            };
            r.readAsArrayBuffer(f)
        };
        document.getElementById('gi-btn').onclick = function () {
            var nk = document.getElementById('gi-col-name').value;
            var sk = document.getElementById('gi-col-grade').value;
            var rows = document.querySelectorAll('#ctl00_ctl00_cphGeneral_cphMain_rgPuntenP_ctl00 tbody tr');
            let m = 0,
                nf = 0;
            var hd = document.getElementById('ctl00_ctl00_cphGeneral_cphMain_txtGewijzigdPagina');
            var gd = window.$find ? window.$find('ctl00_ctl00_cphGeneral_cphMain_rgPuntenP') : null;
            rows.forEach(r => {
                var nc = r.cells[1];
                if (!nc) return;
                var dn = nc.textContent.trim().replace(/\s*\[.*?\]$/, '').trim().toLowerCase();
                var rec = json.find(d => {
                    var en = String(d[nk]).trim().toLowerCase();
                    return dn === en
                });
                if (rec) {
                    var ipt = r.querySelector('input[name*=\'txtScore\']');
                    if (ipt) {
                        ipt.style.backgroundColor = '';
                        nc.style.backgroundColor = '';
                        var v = rec[sk];
                        if (v !== undefined) {
                            var val = String(v).trim();
                            if (val !== '') {
                                if (val.toUpperCase() === 'A') {
                                    ipt.focus();
                                    ipt.dispatchEvent(new KeyboardEvent('keydown', {
                                        key: '+',
                                        code: 'NumpadAdd',
                                        keyCode: 107,
                                        which: 107,
                                        bubbles: true
                                    }));
                                    ipt.dispatchEvent(new KeyboardEvent('keypress', {
                                        key: '+',
                                        keyCode: 43,
                                        which: 43,
                                        bubbles: true
                                    }));
                                    ipt.style.backgroundColor = '#cfe2ff';
                                    m++
                                } else {
                                    if (ipt.onclick) ipt.onclick();
                                    ipt.focus();
                                    if (ipt.onfocus) ipt.onfocus();
                                    ipt.value = val;
                                    if (ipt.onkeydown) ipt.onkeydown({
                                        keyCode: 13
                                    });
                                    if (ipt.onchange) ipt.onchange();
                                    ipt.blur();
                                    if (ipt.onblur) ipt.onblur();
                                    try {
                                        if (hd && gd) {
                                            var ps = r.id.split('__');
                                            if (ps.length > 1) {
                                                var idx = parseInt(ps[ps.length - 1], 10);
                                                var itm = gd.get_masterTableView().get_dataItems()[idx];
                                                if (itm) {
                                                    var id = itm.getDataKeyValue('p_examen');
                                                    if (id && hd.value.indexOf(id + ';') === -1) {
                                                        hd.value += id + ';'
                                                    }
                                                }
                                            }
                                        }
                                    } catch (e) {
                                        console.error(e)
                                    }
                                    m++
                                }
                            }
                        }
                    } else {
                        nf++
                    }
                }
            });
            document.getElementById('gi-msg').innerHTML = '<strong>Done!</strong><br>Matched: ' + m + ' student(s)<br>Not Found: ' + nf + ' student(s)'
        }
    };
    document.head.appendChild(s)
}
