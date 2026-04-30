#!/usr/bin/env python3
"""
business_management_workflow_v3.json の修正スクリプト
修正内容:
1. LINEメッセージ解析: ファイル名 経営管理→経費管理、月ゼロパディング
2. Excel書き込み: 経費ヘッダー8列化・確認フラグ削除・備考に統合
3. Excel書き込み: 外注管理・売上管理ヘッダーの括弧を全角に統一
4. Excel書き込み: 月次サマリー更新を直接行番号指定に変更（粗利はE37）
"""
import json
import re

SRC = '/home/user/genai-lessons/workflows/business_management_workflow_v3.json'
DST = '/home/user/genai-lessons/workflows/business_management_workflow_v3.json'

with open(SRC, encoding='utf-8') as f:
    w = json.load(f)

# ============================================================
# Fix 1: LINEメッセージ解析 - ファイル名修正
# ============================================================
for node in w['nodes']:
    if node['name'] == 'LINEメッセージ解析':
        code = node['parameters']['jsCode']
        # 経営管理_ → 経費管理_
        code = code.replace('経営管理_', '経費管理_')
        # 月のゼロパディング: ${mo}月 → ${String(mo).padStart(2,'0')}月
        code = code.replace("${mo}月", "${String(mo).padStart(2,'0')}月")
        # ${tm}月 → ${String(tm).padStart(2,'0')}月
        code = code.replace("${tm}月", "${String(tm).padStart(2,'0')}月")
        node['parameters']['jsCode'] = code
        print('✓ LINEメッセージ解析: ファイル名修正完了')

# ============================================================
# Fix 2 & 3 & 4: Excel書き込み＋アップロード
# ============================================================
NEW_EXCEL_CODE = r"""
const ExcelJS = require('exceljs');
return await (async () => {
  const data = $('LINEメッセージ解析').first().json;
  const fileInfo = $('月次ファイル情報取得').first().json;
  try {
    const wb = new ExcelJS.Workbook();
    const dlUrl = fileInfo['@microsoft.graph.downloadUrl'];
    if (dlUrl) {
      try {
        const r = await this.helpers.httpRequest({ method:'GET', url:dlUrl, encoding:'arraybuffer', returnFullResponse:false });
        await wb.xlsx.load(Buffer.from(r), { useSharedStrings: false });
      } catch(e) {}
    }
    const toDate = s => { if(!s)return null; const p=s.split('/'); return new Date(Date.UTC(+p[0],+p[1]-1,+p[2])); };
    const PAYMENT_RULES = {
      'オークコミュニケーション': { months: 2, day: 10 },
      '平成システム': { cutoffDay: 20 }
    };
    const calcPaymentDue = (regDateStr, clientName) => {
      const rd = toDate(regDateStr);
      if (!rd) return null;
      const rule = PAYMENT_RULES[clientName];
      if (rule) {
        if (rule.cutoffDay) {
          const add = rd.getDate() <= rule.cutoffDay ? 1 : 2;
          return new Date(Date.UTC(rd.getFullYear(), rd.getMonth() + add + 1, 0));
        }
        const dm = rd.getMonth() + rule.months;
        return new Date(Date.UTC(rd.getFullYear() + Math.floor(dm/12), dm%12, rule.day));
      }
      return new Date(Date.UTC(rd.getFullYear(), rd.getMonth()+2, 0));
    };
    const getV = c => (c&&typeof c==='object'&&'result'in c)?c.result:c;
    const SUBCON_RULES = {
      '秀電工': { months: 2, day: 10 }
    };
    const calcSubconDue = (regDateStr, vendorName) => {
      const rd = toDate(regDateStr);
      if (!rd) return null;
      const rule = SUBCON_RULES[vendorName];
      if (rule) {
        const dm = rd.getMonth() + rule.months;
        return new Date(Date.UTC(rd.getFullYear() + Math.floor(dm/12), dm%12, rule.day));
      }
      return new Date(Date.UTC(rd.getFullYear(), rd.getMonth()+2, 0));
    };
    const ensure = (name, hdrs) => {
      let s = wb.getWorksheet(name);
      if (!s) { s=wb.addWorksheet(name); const h=s.getRow(1); hdrs.forEach((v,i)=>h.getCell(i+1).value=v); h.font={bold:true}; h.commit(); }
      return s;
    };
    const lastR = (s, col) => { let r=1; s.eachRow((row,n)=>{ if(n>=2&&getV(row.getCell(col).value)!=null&&getV(row.getCell(col).value)!=='')r=n; }); return r; };
    const mt = data.messageType;

    // ===== 経費 =====
    // A=S(連番) B=日付 C=経費種類 D=支払方法 E=店名・支払先 F=案件名 G=金額（税込） H=備考
    if (mt === 'expense') {
      const s = ensure('経費', ['S','日付','経費種類','支払方法','店名・支払先','案件名','金額（税込）','備考']);
      const nr = lastR(s,2)+1; const row = s.getRow(nr);
      row.getCell(1).value = nr-1;
      row.getCell(2).value = toDate(data.date); row.getCell(2).numFmt='YYYY/MM/DD';
      row.getCell(3).value = data.expenseCategory;
      row.getCell(4).value = data.paymentMethod;
      row.getCell(5).value = data.storeName;
      row.getCell(6).value = data.projectName||'';
      row.getCell(7).value = data.amount; row.getCell(7).numFmt='#,##0';
      row.getCell(8).value = data.requiresReview ? `⚠要確認 ${data.note||''}`.trim() : (data.note||'');
      row.commit();

    } else if (mt === 'multi_expense') {
      const s = ensure('経費', ['S','日付','経費種類','支払方法','店名・支払先','案件名','金額（税込）','備考']);
      const entries = data.entries || [];
      let totalAmt = 0; let hasReview = false;
      for (const entry of entries) {
        const nr = lastR(s,2)+1; const row = s.getRow(nr);
        const amt = entry.amount || 0;
        row.getCell(1).value = nr-1;
        row.getCell(2).value = toDate(entry.date); row.getCell(2).numFmt='YYYY/MM/DD';
        row.getCell(3).value = entry.expenseCategory||'雑費';
        row.getCell(4).value = entry.paymentMethod||'口座引落';
        row.getCell(5).value = entry.storeName||'';
        row.getCell(6).value = '';
        row.getCell(7).value = amt; row.getCell(7).numFmt='#,##0';
        row.getCell(8).value = entry.requiresReview ? '⚠要確認' : '';
        row.commit();
        totalAmt += amt;
        if (entry.requiresReview) hasReview = true;
      }
      data.entryCount = entries.length;
      data.totalAmount = totalAmt;
      data.hasReview = hasReview;

    // ===== 外注管理 =====
    // A=No. B=登録日 C=業者名 D=種別 E=インボイス F=案件名 G=請求額（税抜） H=消費税（10%） I=合計（税込） J=請求日 K=支払期限 L=支払日 M=支払額 N=状況
    } else if (mt === 'subcontractor_invoice') {
      const s = ensure('外注管理', ['No.','登録日','業者名','種別','インボイス','案件名','請求額（税抜）','消費税（10%）','合計（税込）','請求日','支払期限','支払日','支払額','状況']);
      const nr = lastR(s,2)+1; const row = s.getRow(nr);
      const exTax = data.amountExTax || 0;
      const tax = data.tax !== undefined ? data.tax : Math.floor(exTax * 0.1);
      const total = data.totalAmount || (exTax + tax);
      const due = calcSubconDue(data.registrationDate, data.vendorName);
      const dueStr = due ? `${due.getUTCFullYear()}/${String(due.getUTCMonth()+1).padStart(2,'0')}/${String(due.getUTCDate()).padStart(2,'0')}` : '';
      row.getCell(1).value = nr-1;
      row.getCell(2).value = toDate(data.registrationDate); row.getCell(2).numFmt='YYYY/MM/DD';
      row.getCell(3).value = data.vendorName;
      row.getCell(4).value = data.vendorType||'';
      row.getCell(5).value = data.invoice||'';
      row.getCell(6).value = data.projectName||'';
      row.getCell(7).value = exTax; row.getCell(7).numFmt='#,##0';
      row.getCell(8).value = tax; row.getCell(8).numFmt='#,##0';
      row.getCell(9).value = total; row.getCell(9).numFmt='#,##0';
      row.getCell(10).value = toDate(data.registrationDate); row.getCell(10).numFmt='YYYY/MM/DD';
      if(due){ row.getCell(11).value=due; row.getCell(11).numFmt='YYYY/MM/DD'; }
      row.getCell(14).value = '未払';
      row.commit();
      data.calculatedDue = dueStr; data.totalAmount = total; data.tax = tax; data.amountExTax = exTax;

    } else if (mt === 'subcontractor_payment') {
      const s = ensure('外注管理', ['No.','登録日','業者名','種別','インボイス','案件名','請求額（税抜）','消費税（10%）','合計（税込）','請求日','支払期限','支払日','支払額','状況']);
      let target = null;
      s.eachRow((row,n)=>{ if(n<2||target)return; if(getV(row.getCell(3).value)===data.vendorName&&getV(row.getCell(14).value)==='未払')target=row; });
      if (!target) return [{ json: { valid:false, replyToken:data.replyToken, errorMessage:`${data.vendorName}の未払い請求が見つかりません` } }];
      target.getCell(12).value = toDate(data.paymentDate); target.getCell(12).numFmt='YYYY/MM/DD';
      target.getCell(13).value = data.paymentAmount; target.getCell(13).numFmt='#,##0';
      target.getCell(14).value = '支払済';
      target.commit();

    // ===== 売上管理 =====
    // A=No. B=登録日 C=得意先名 D=案件名 E=請求額（税抜） F=消費税（10%） G=合計（税込） H=請求日 I=入金期限 J=入金日 K=入金額 L=状況
    } else if (mt === 'sales') {
      const s = ensure('売上管理', ['No.','登録日','得意先名','案件名','請求額（税抜）','消費税（10%）','合計（税込）','請求日','入金期限','入金日','入金額','状況']);
      const nr = lastR(s,2)+1; const row = s.getRow(nr);
      row.getCell(1).value = nr-1;
      row.getCell(2).value = toDate(data.registrationDate); row.getCell(2).numFmt='YYYY/MM/DD';
      row.getCell(3).value = data.clientName;
      row.getCell(4).value = data.projectName||'';
      row.getCell(5).value = data.amountExTax; row.getCell(5).numFmt='#,##0';
      row.getCell(6).value = data.tax; row.getCell(6).numFmt='#,##0';
      row.getCell(7).value = data.totalAmount; row.getCell(7).numFmt='#,##0';
      row.getCell(8).value = toDate(data.invoiceDate); row.getCell(8).numFmt='YYYY/MM/DD';
      const dueDate = data.paymentDueDate ? toDate(data.paymentDueDate) : calcPaymentDue(data.registrationDate, data.clientName);
      if(dueDate){ row.getCell(9).value=dueDate; row.getCell(9).numFmt='YYYY/MM/DD'; }
      row.getCell(12).value = '未入金';
      row.commit();

    } else if (mt === 'payment_received') {
      const s = ensure('売上管理', ['No.','登録日','得意先名','案件名','請求額（税抜）','消費税（10%）','合計（税込）','請求日','入金期限','入金日','入金額','状況']);
      let target = null;
      s.eachRow((row,n)=>{ if(n<2||target)return; if(getV(row.getCell(3).value)===data.clientName&&getV(row.getCell(12).value)==='未入金')target=row; });
      if (!target) return [{ json: { valid:false, replyToken:data.replyToken, errorMessage:`${data.clientName}の未入金請求が見つかりません` } }];
      target.getCell(10).value = toDate(data.receivedDate); target.getCell(10).numFmt='YYYY/MM/DD';
      target.getCell(11).value = data.receivedAmount; target.getCell(11).numFmt='#,##0';
      target.getCell(12).value = '入金済';
      target.commit();
    }

    // ===== 月次サマリー更新（直接行番号指定）=====
    // Row4=請求済合計（税込）B  Row5=入金済合計 B  Row6=未入金残高 B
    // Row9=請求受取合計（税込）B  Row10=支払済合計 B  Row11=未払残高 B
    // Row14-29=経費カテゴリ別 B  Row30=経費合計 B
    // Row33=支給総額合計 B  Row34=手取り合計 B（人件費シートから読む）
    // Row37=粗利（概算）E
    const summary = wb.getWorksheet('月次サマリー');
    if (summary) {
      const getV2 = c => (c&&typeof c==='object'&&'result'in c)?c.result:c;
      const expS=wb.getWorksheet('経費'), subS=wb.getWorksheet('外注管理'), salS=wb.getWorksheet('売上管理'), payS=wb.getWorksheet('人件費');

      const expCats=['材料費','燃料費','高速料金','駐車場代','工具・消耗品費','通信費','交際費','会議費','事務用品費','広告宣伝費','研修費','福利厚生費','車両維持費','地代家賃','保険料','雑費'];
      const byCat={};
      if(expS) expS.eachRow((r,n)=>{ if(n<2)return; const c=getV2(r.getCell(3).value); const a=Number(getV2(r.getCell(7).value))||0; if(c) byCat[c]=(byCat[c]||0)+a; });
      const expTotal = Object.values(byCat).reduce((a,b)=>a+b, 0);

      let salTotal=0, recvTotal=0;
      if(salS) salS.eachRow((r,n)=>{ if(n<2)return; const t=Number(getV2(r.getCell(7).value))||0; const st=getV2(r.getCell(12).value); salTotal+=t; if(st==='入金済') recvTotal+=Number(getV2(r.getCell(11).value))||t; });

      let subTotal=0, subPaid=0;
      if(subS) subS.eachRow((r,n)=>{ if(n<2)return; const t=Number(getV2(r.getCell(9).value))||0; const st=getV2(r.getCell(14).value); subTotal+=t; if(st==='支払済') subPaid+=Number(getV2(r.getCell(13).value))||t; });

      let payTotal=0, takeHomeTotal=0;
      if(payS) payS.eachRow((r,n)=>{ if(n<2||n>14)return; payTotal+=Number(getV2(r.getCell(2).value))||0; takeHomeTotal+=Number(getV2(r.getCell(4).value))||0; });

      const setSum = (row, col, val) => {
        const c = summary.getCell(row, col);
        c.value = val;
        c.numFmt = '#,##0';
        summary.getRow(row).commit();
      };

      setSum(4, 2, salTotal);
      setSum(5, 2, recvTotal);
      setSum(6, 2, salTotal - recvTotal);
      setSum(9, 2, subTotal);
      setSum(10, 2, subPaid);
      setSum(11, 2, subTotal - subPaid);
      expCats.forEach((cat, i) => setSum(14+i, 2, byCat[cat]||0));
      setSum(30, 2, expTotal);
      setSum(33, 2, payTotal);
      setSum(34, 2, takeHomeTotal);
      setSum(37, 5, recvTotal - subPaid - expTotal - payTotal);
    }

    // OneDriveアップロード
    const creds = await this.getCredentials('microsoftDriveOAuth2Api');
    const token = creds.oauthTokenData?.access_token;
    if (!token) throw new Error('Microsoft認証トークンが取得できません。資格情報を再接続してください。');
    const buf = await wb.xlsx.writeBuffer({ useSharedStrings: false });
    const upUrl = fileInfo.id
      ? `https://graph.microsoft.com/v1.0/me/drive/items/${fileInfo.id}/content`
      : `https://graph.microsoft.com/v1.0/me/drive/root:/業務管理システム/経費管理/月次/${data.fileName}:/content`;
    await this.helpers.httpRequest({
      method:'PUT', url:upUrl,
      headers:{ 'Authorization':`Bearer ${token}`, 'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      body:Buffer.from(buf), returnFullResponse:false
    });
    return [{ json: { ...data, valid: true } }];
  } catch(e) {
    return [{ json: { valid:false, replyToken:data.replyToken, errorMessage:e.message } }];
  }
})();
""".strip()

for node in w['nodes']:
    if node['name'] == 'Excel書き込み＋アップロード':
        node['parameters']['jsCode'] = NEW_EXCEL_CODE
        print('✓ Excel書き込み＋アップロード: コード修正完了')

with open(DST, 'w', encoding='utf-8') as f:
    json.dump(w, f, ensure_ascii=False, indent=2)

print(f'✓ 保存完了: {DST}')
