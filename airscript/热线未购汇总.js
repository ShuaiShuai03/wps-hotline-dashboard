/* ============================================================
   热线 & 未购｜看板数据导出（AirScript · 只读）
   放在任意一个战区管理系统的脚本编辑器里直接运行。
   - 按【字段名】读取，字段顺序被调整不影响结果
   - 只读源表：热线日报表 / 未购日报表，不依赖任何统计表、合并表
   - 运行后把控制台打印的 JSON 整段复制，粘贴进网页看板
   ============================================================ */

// 表名可按战区改；脚本会自动匹配"包含关键字"的表，通常无需改动
var HOTLINE_KEY = '热线日报';
var NOBUY_KEY   = '未购日报';
var ZONE        = '';   // 留空则自动从表名推断，如"一战区热线日报表" -> 一战区

var HOTLINE_FIELDS = ['轮次','部门','组别','员工','引流渠道','创建时间','广告费','进线量','注册',
  '当日完播','采集总','定金','实发单量','实发金额','签收金额','线上产出','线上单量','战区'];
var NOBUY_FIELDS = ['轮次','部门','组别','员工','引流渠道','客户类型','创建时间','线索成本','承接数据',
  '首轮完播','首轮单量','首轮发货业绩','往期完播','往期单量','往期发货业绩',
  '总成交单量','总发货业绩','总活跃客户','总待激活','流失','签收业绩'];

function findSheetByKey(key){
  var n = Application.Sheets.Count;
  for (var i = 1; i <= n; i++){
    var s = Application.Sheets.Item(i);
    if (s.Name && s.Name.indexOf(key) >= 0 && s.Name.indexOf('统计') < 0 && s.Name.indexOf('📊') < 0) return s;
  }
  return null;
}

function fieldMap(sheet){
  var m = {};
  sheet.GetFields().forEach(function(f){ m[f.name] = f.id; });
  return m;
}

// 单元格取值：兼容 人员/关联/选项 等对象型字段
function cellValue(sheet, row, fmap, name){
  var fid = fmap[name];
  if (!fid) return null;
  try {
    var v = sheet.RecordRange(row, fid).Value;
    if (v && typeof v === 'object'){
      if (Array.isArray(v)) {
        return v.map(function(p){ return (p && (p.nickname || p.name || p.text)) || p; }).join(',');
      }
      if (Array.isArray(v.Value)) {
        return v.Value.map(function(p){ return (p && (p.nickname || p.name || p.text || p.userId)) || p; }).join(',');
      }
      if (v.Value !== undefined) return v.Value;
      return null;
    }
    return v;
  } catch (e) { return null; }
}

function readRows(sheet, fields, tag, zone){
  var fmap = fieldMap(sheet);
  var total = sheet.RecordRange().Count;
  var out = [];
  for (var r = 1; r <= total; r++){
    var o = { 类型: tag, 战区: zone };
    var hasData = false;
    for (var i = 0; i < fields.length; i++){
      var name = fields[i];
      var v = cellValue(sheet, r, fmap, name);
      if (name === '战区'){ if (v) o.战区 = v; continue; }
      o[name] = v;
      if (v !== null && v !== '' && v !== 0) hasData = true;
    }
    if (hasData) out.push(o);
  }
  return out;
}

function guessZone(name){
  var m = String(name || '').match(/(.{1,4}战区)/);
  return m ? m[1] : '';
}

function main(){
  var hlSheet = findSheetByKey(HOTLINE_KEY);
  var nbSheet = findSheetByKey(NOBUY_KEY);
  var zone = ZONE || guessZone(hlSheet && hlSheet.Name) || guessZone(nbSheet && nbSheet.Name) || '本战区';

  var detail = [];
  var srcs = [];
  if (hlSheet){ var hl = readRows(hlSheet, HOTLINE_FIELDS, '热线', zone); detail = detail.concat(hl); srcs.push(hlSheet.Name + '(' + hl.length + '条)'); }
  else console.log('⚠️ 未找到包含「' + HOTLINE_KEY + '」的数据表');
  if (nbSheet){ var nb = readRows(nbSheet, NOBUY_FIELDS, '未购', zone); detail = detail.concat(nb); srcs.push(nbSheet.Name + '(' + nb.length + '条)'); }
  else console.log('⚠️ 未找到包含「' + NOBUY_KEY + '」的数据表');

  var payload = {
    meta: { zone: zone, generatedAt: new Date().toISOString(), source: srcs.join(' + ') },
    detail: detail
  };
  console.log('==== 复制下面这一整段 JSON ====');
  console.log(JSON.stringify(payload));
  console.log('==== 共 ' + detail.length + ' 条记录 ====');
  return payload;
}

main();
