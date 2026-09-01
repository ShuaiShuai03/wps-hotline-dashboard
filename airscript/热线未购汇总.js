/* ============================================================
   热线 & 未购｜看板数据导出（AirScript · 只读）
   放在任意一个战区管理系统的脚本编辑器里直接运行。
   - 按【字段名】读取，字段顺序被调整不影响结果
   - 只读源表：热线日报表 / 未购日报表，不依赖任何统计表、合并表
   - 输出分段打印（每段 700 字），因为脚本日志面板会把超长的单行截断。
     运行完在日志区全选复制（Cmd/Ctrl+A → Cmd/Ctrl+C），
     整段粘进看板的「导入 AirScript 数据」即可，看板会自动拼接。
   ============================================================ */

// 表名按"包含关键字"匹配，通常无需改动
var HOTLINE_KEY = '热线日报';
var NOBUY_KEY   = '未购日报';
// ZONE 留空则自动从表名推断，如"一战区热线日报表" -> 一战区
var ZONE = '';

var HOTLINE_FIELDS = ['轮次','部门','组别','员工','引流渠道','创建时间','广告费','进线量','注册',
  '当日完播','采集总','定金','实发单量','实发金额','签收金额','线上产出','线上单量'];
var NOBUY_FIELDS = ['轮次','部门','组别','员工','引流渠道','客户类型','创建时间','线索成本','承接数据',
  '首轮完播','首轮单量','首轮发货业绩','往期完播','往期单量','往期发货业绩',
  '总成交单量','总发货业绩','总活跃客户','总待激活','流失','签收业绩'];

var CHUNK = 700;

// 返回所有匹配的表：单战区文档只有一张，总管理系统里每个战区各一张，全部导出
function findSheetsByKey(key){
  var n = Application.Sheets.Count, out = [];
  for (var i = 1; i <= n; i++){
    var s = Application.Sheets.Item(i);
    if (s.Name && s.Name.indexOf(key) >= 0 && s.Name.indexOf('统计') < 0 && s.Name.indexOf('📊') < 0) out.push(s);
  }
  return out;
}

// ⚠️ 本脚本要求脚本环境 AirScript 2.0（新建脚本时在 + 右侧的下拉里选）。
// 在 1.0 环境下 GetFields() 返回的 f.name 是 Promise，会变成 "[object Promise]"，
// 导致字段名->id 的映射整体错位、取到完全错误的列；而 1.0 的沙箱又缺少可用的
// Promise.all/async 支持，无法在脚本内绕开。所以统一用 2.0。
function fieldMap(sheet){
  var m = {};
  sheet.GetFields().forEach(function(f){ m[String(f.name)] = f.id; });
  return m;
}

// 源表缺字段时明确报出来，而不是静默返回空列
function warnMissing(sheetName, fmap, fields){
  var miss = fields.filter(function(n){ return !fmap[n]; });
  if (miss.length) console.log('⚠️ ' + sheetName + ' 缺少字段：' + miss.join('、'));
}

// 关联/人员/选项字段的单元格是 DBCellValue：{ Value:[{id,str}], display, ... }
// 取 str（关联/人员显示名），其次 nickname/name/text，最后 display。
function textOf(p){
  if (p == null) return '';
  if (typeof p !== 'object') {
    var t = String(p);
    // 兜底：任何被字符串化的对象（[object Promise] / [object Object]）都当空值丢掉
    return /^\[object /.test(t) ? '' : t;
  }
  var s = p.str || p.nickname || p.name || p.text || p.title || '';
  s = String(s);
  return /^\[object /.test(s) ? '' : s;
}

// 单元格取值：两个脚本环境下都是同步的
//   2.0 -> DBCellValue 对象 { Value:[{id,str}], display }
//   1.0 -> 普通数组 ["热线二部"] / 字符串 / 数字
function cellValue(sheet, row, fmap, name){
  var fid = fmap[name];
  if (!fid) return null;
  try {
    var v = sheet.RecordRange(row, fid).Value;
    if (v == null) return null;
    if (typeof v !== 'object') return v;
    if (Array.isArray(v)){
      return v.map(textOf).filter(function(x){ return x !== ''; }).join(',');
    }
    if (Array.isArray(v.Value)){
      var joined = v.Value.map(textOf).filter(function(x){ return x !== ''; }).join(',');
      return joined || (v.display || '');
    }
    if (v.Value !== undefined && typeof v.Value !== 'object') return v.Value;
    return v.display || '';
  } catch (e) { return null; }
}

function readRows(sheet, fmap, fields, tag, zone, cols){
  var total = sheet.RecordRange().Count;
  var out = [];
  for (var r = 1; r <= total; r++){
    var rec = [], hasData = false;
    for (var c = 0; c < cols.length; c++){
      var name = cols[c], v;
      if (name === '类型') v = tag;
      else if (name === '战区') v = zone;
      else if (fields.indexOf(name) < 0) v = null;
      else {
        v = cellValue(sheet, r, fmap, name);
        // 创建时间是系统自动填的，不能当作"这一行有数据"的依据，否则空行也会被导出
        if (name !== '创建时间' && v !== null && v !== '' && v !== 0) hasData = true;
      }
      rec.push(v === undefined ? null : v);
    }
    if (hasData) out.push(rec);
  }
  return out;
}

function guessZone(name){
  var m = String(name || '').match(/(.{1,4}战区)/);
  return m ? m[1] : '';
}

// 分段打印：日志面板会截断超长单行，所以切成小段，看板端自动拼接
function emit(str){
  var total = Math.ceil(str.length / CHUNK);
  console.log('==== 数据开始：共 ' + total + ' 段，请在日志区全选复制后粘进看板 ====');
  for (var i = 0; i < total; i++){
    console.log('@@' + (i + 1) + '@@' + str.substr(i * CHUNK, CHUNK) + '@@/@@');
  }
  console.log('==== 数据结束 ====');
}

function main(){
  var hlSheets = findSheetsByKey(HOTLINE_KEY);
  var nbSheets = findSheetsByKey(NOBUY_KEY);
  if (!hlSheets.length) console.log('未找到包含「' + HOTLINE_KEY + '」的数据表');
  if (!nbSheets.length) console.log('未找到包含「' + NOBUY_KEY + '」的数据表');

  // 列式输出：所有行共用一份表头，体积比对象数组小一半以上
  var cols = ['类型','战区'];
  HOTLINE_FIELDS.concat(NOBUY_FIELDS).forEach(function(f){ if (cols.indexOf(f) < 0) cols.push(f); });

  var tasks = [];
  hlSheets.forEach(function(sh){ tasks.push({ sh: sh, fields: HOTLINE_FIELDS, tag: '热线' }); });
  nbSheets.forEach(function(sh){ tasks.push({ sh: sh, fields: NOBUY_FIELDS, tag: '未购' }); });

  var rows = [], srcs = [], zones = [];
  tasks.forEach(function(t){
    var z = ZONE || guessZone(t.sh.Name) || '本战区';
    if (zones.indexOf(z) < 0) zones.push(z);
    var fmap = fieldMap(t.sh);
    warnMissing(t.sh.Name, fmap, t.fields);
    var got = readRows(t.sh, fmap, t.fields, t.tag, z, cols);
    rows = rows.concat(got);
    srcs.push(t.sh.Name + '(' + got.length + '条)');
  });

  var payload = {
    meta: { zone: zones.join('、') || '本战区', generatedAt: new Date().toISOString(), source: srcs.join(' + ') },
    cols: cols,
    rows: rows
  };
  emit(JSON.stringify(payload));
  console.log('共 ' + rows.length + ' 条记录，涵盖战区：' + zones.join('、'));
  return payload;
}

main();
