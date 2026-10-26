/**
 * ملعب جمعية الراكة — Google Apps Script API
 * التشغيل الأول: شغّل setup() يدويًا من المحرر ووافق على الصلاحيات.
 */
const CONFIG={SPREADSHEET_ID:'',RECEIPTS_FOLDER_ID:'',SESSION_TTL_SECONDS:21600,MAX_RECEIPT_BYTES:4*1024*1024,ALLOWED_TEAMS:['سعودي','مصري','سوري','فلبيني','هندي','أكاديمية'],ALLOWED_ROLES:['admin','viewer','employee']};
const USERS='المستخدمون', REVENUES='الإيرادات', SETTINGS='الإعدادات';

function setup(){
 const p=PropertiesService.getScriptProperties();let ss;const saved=CONFIG.SPREADSHEET_ID||p.getProperty('SPREADSHEET_ID');
 if(saved)ss=SpreadsheetApp.openById(saved);else{ss=SpreadsheetApp.create('ملعب جمعية الراكة - الإيرادات');p.setProperty('SPREADSHEET_ID',ss.getId())}
 ensureSheet_(ss,USERS,['userId','name','username','salt','passwordHash','role','active','createdAt']);
 ensureSheet_(ss,REVENUES,['id','timestamp','team','amount','paymentMethod','createdBy','createdByUsername','receiptFileId','receiptName','receiptMime']);
 ensureSheet_(ss,SETTINGS,['key','value']);
 let folderId=CONFIG.RECEIPTS_FOLDER_ID||p.getProperty('RECEIPTS_FOLDER_ID');
 if(!folderId){folderId=DriveApp.createFolder('ملعب جمعية الراكة - صور الإيصالات').getId();p.setProperty('RECEIPTS_FOLDER_ID',folderId)}
 DriveApp.getFolderById(folderId).setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.NONE);
 p.setProperty('SPREADSHEET_ID',ss.getId());p.setProperty('RECEIPTS_FOLDER_ID',folderId);
 const sh=ss.getSheetByName(USERS);
 if(sh.getLastRow()<=1){const password=Utilities.getUuid().replace(/-/g,'')+'A9!',salt=Utilities.getUuid();sh.appendRow([Utilities.getUuid(),'المدير العام','admin',salt,hashPassword_(password,salt),'admin',true,new Date().toISOString()]);p.setProperty('INITIAL_ADMIN_PASSWORD',password);Logger.log('اسم المستخدم الأول: admin');Logger.log('كلمة مرور المدير المؤقتة (احفظها): '+password)}
 Logger.log('رابط Google Sheets: '+ss.getUrl());Logger.log('معرّف مجلد الإيصالات: '+folderId);return {spreadsheetUrl:ss.getUrl(),folderId:folderId};
}
function doGet(){return out_({ok:true,service:'Rakah Stadium Revenue API',message:'API is running. Send POST requests.'})}
function doPost(e){
 try{const b=e&&e.postData&&e.postData.contents?JSON.parse(e.postData.contents):{},a=String(b.action||'');
 if(a==='login')return out_(login_(b));if(a==='logout')return out_(logout_(b));
 const u=auth_(b.token);
 if(a==='listRevenues')return out_(listRevenues_(u));
 if(a==='deleteRevenue')return out_(deleteRevenue_(u,b));
 if(a==='getReceipt')return out_(getReceipt_(u,b));
 if(a==='addRevenue')return out_(addRevenue_(u,b));
 if(a==='createUser')return out_(createUser_(u,b));
 if(a==='listUsers')return out_(listUsers_(u));
 throw Error('طلب غير معروف.');
 }catch(err){return out_({ok:false,error:safeError_(err)})}
}
function login_(b){
 const username=String(b.username||'').trim().toLowerCase(),password=String(b.password||'');
 if(!username||!password)throw Error('أدخل اسم المستخدم وكلمة المرور.');
 const v=sheet_(USERS).getDataRange().getValues();
 for(let i=1;i<v.length;i++){const r=v[i];if(String(r[2]).toLowerCase()!==username)continue;
 if(!truth_(r[6]))throw Error('هذا الحساب موقوف.');
 if(hashPassword_(password,String(r[3]))!==String(r[4]))throw Error('اسم المستخدم أو كلمة المرور غير صحيحة.');
 const token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
 const user={id:String(r[0]),name:String(r[1]),username:String(r[2]),role:String(r[5])};
 CacheService.getScriptCache().put('session:'+token,JSON.stringify(user),CONFIG.SESSION_TTL_SECONDS);return {ok:true,token:token,user:user}
 }throw Error('اسم المستخدم أو كلمة المرور غير صحيحة.')
}
function logout_(b){const t=String(b.token||'');if(t)CacheService.getScriptCache().remove('session:'+t);return {ok:true}}
function auth_(token){
 if(!token||typeof token!=='string'||token.length>200)throw Error('انتهت الجلسة؛ سجّل الدخول مجددًا.');
 const raw=CacheService.getScriptCache().get('session:'+token);if(!raw)throw Error('انتهت الجلسة؛ سجّل الدخول مجددًا.');
 const s=JSON.parse(raw),v=sheet_(USERS).getDataRange().getValues();
 for(let i=1;i<v.length;i++){const r=v[i];if(String(r[0])===s.id){if(!truth_(r[6]))throw Error('الحساب موقوف.');s.role=String(r[5]);s.name=String(r[1]);s.username=String(r[2]);return s}}
 throw Error('الحساب غير موجود.')
}
function listRevenues_(u){
 requireRole_(u,['admin','viewer','employee']);const v=sheet_(REVENUES).getDataRange().getValues(),out=[];
 for(let i=1;i<v.length;i++){
  const r=v[i];if(!r[0])continue;
  // موظف إدخال البيانات يرى عملياته فقط؛ المدير والمشاهد يحتفظان بنطاق العرض الحالي.
  if(u.role==='employee'&&String(r[6]||'').trim().toLowerCase()!==String(u.username||'').trim().toLowerCase())continue;
  const d=r[1] instanceof Date?r[1]:new Date(r[1]);
  out.push({id:String(r[0]),timestamp:isNaN(d.getTime())?'':Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd'),timestampDisplay:isNaN(d.getTime())?'':Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy/MM/dd HH:mm'),team:String(r[2]),amount:Number(r[3]),paymentMethod:String(r[4]),createdBy:String(r[5])});
 }
 out.sort((a,b)=>b.timestampDisplay.localeCompare(a.timestampDisplay));return {ok:true,revenues:out}
}
function addRevenue_(u,b){
 requireRole_(u,['admin','employee']);const team=String(b.team||''),amount=Number(b.amount),payment=String(b.paymentMethod||'شبكة');
 if(CONFIG.ALLOWED_TEAMS.indexOf(team)<0)throw Error('الفريق المحدد غير صحيح.');
 if(!isFinite(amount)||amount<=0||amount>1000000)throw Error('المبلغ يجب أن يكون أكبر من صفر وأقل من مليون ريال.');
 if(payment!=='شبكة')throw Error('طريقة الدفع المسموحة هي الشبكة فقط.');
 const data=String(b.receiptData||''),m=data.match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);if(!m)throw Error('أرفق صورة JPG أو PNG أو WEBP.');
 const bytes=Utilities.base64Decode(m[2]);if(bytes.length>CONFIG.MAX_RECEIPT_BYTES)throw Error('حجم الصورة أكبر من 4 ميجابايت.');
 const name=String(b.receiptName||'receipt').replace(/[^\w.\-\u0600-\u06FF ]/g,'_').slice(0,100),blob=Utilities.newBlob(bytes,m[1],name||'receipt');
 const folder=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('RECEIPTS_FOLDER_ID')),file=folder.createFile(blob);file.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.NONE);
 let now=new Date();if(b.entryDateTime){if(u.role!=='admin')throw Error('تعديل تاريخ الإيراد متاح للمدير العام فقط.');const requested=new Date(String(b.entryDateTime));if(isNaN(requested.getTime()))throw Error('تاريخ ووقت الإيراد غير صحيح.');if(requested.getTime()>Date.now())throw Error('لا يمكن تسجيل إيراد بتاريخ مستقبلي.');now=requested;}const id=Utilities.getUuid(),lock=LockService.getScriptLock();lock.waitLock(20000);
 try{sheet_(REVENUES).appendRow([id,now,team,amount,payment,u.name,u.username,file.getId(),file.getName(),m[1]])}finally{lock.releaseLock()}
 return {ok:true,id:id,message:'تم تسجيل الإيراد.'}
}
function deleteRevenue_(u,b){
 requireRole_(u,['admin']);
 const id=String(b.id||'').trim();if(!id)throw Error('رقم العملية غير صحيح.');
 const sh=sheet_(REVENUES),values=sh.getDataRange().getValues();let row=-1,fileId='';
 for(let i=1;i<values.length;i++){if(String(values[i][0])===id){row=i+1;fileId=String(values[i][7]||'');break}}
 if(row<0)throw Error('العملية غير موجودة أو تم حذفها مسبقًا.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  // Re-check under lock so a stale page cannot delete a different record.
  const current=sh.getDataRange().getValues();let currentRow=-1,currentFile='';
  for(let i=1;i<current.length;i++){if(String(current[i][0])===id){currentRow=i+1;currentFile=String(current[i][7]||'');break}}
  if(currentRow<0)throw Error('العملية غير موجودة أو تم حذفها مسبقًا.');
  if(currentFile){try{DriveApp.getFileById(currentFile).setTrashed(true)}catch(fileErr){throw Error('تعذر نقل صورة الإيصال إلى المهملات؛ لم يتم حذف العملية.')}}
  sh.deleteRow(currentRow);
 }finally{lock.releaseLock()}
 return {ok:true,message:'تم حذف العملية وصورة الإيصال.'}
}
function getReceipt_(u,b){
 requireRole_(u,['admin','viewer','employee']);const id=String(b.id||'');if(!id)throw Error('رقم الإيصال غير صحيح.');
 const v=sheet_(REVENUES).getDataRange().getValues();let fid='',owner='';
 for(let i=1;i<v.length;i++)if(String(v[i][0])===id){fid=String(v[i][7]);owner=String(v[i][6]||'').trim().toLowerCase();break}
 if(!fid)throw Error('الإيصال غير موجود.');
 if(u.role==='employee'&&owner!==String(u.username||'').trim().toLowerCase())throw Error('ليس لديك صلاحية عرض هذا الإيصال.');
 const blob=DriveApp.getFileById(fid).getBlob(),bytes=blob.getBytes();
 if(bytes.length>CONFIG.MAX_RECEIPT_BYTES)throw Error('تعذر عرض الإيصال بسبب الحجم.');
 return {ok:true,dataUrl:'data:'+blob.getContentType()+';base64,'+Utilities.base64Encode(bytes)}
}
function createUser_(u,b){
 requireRole_(u,['admin']);const name=String(b.name||'').trim(),username=String(b.username||'').trim().toLowerCase(),password=String(b.password||''),role=String(b.role||'');
 if(name.length<2||name.length>80)throw Error('اكتب اسمًا صحيحًا.');
 if(!/^[a-z0-9._-]{3,40}$/.test(username))throw Error('اسم المستخدم يجب أن يكون 3–40 حرفًا إنجليزيًا أو رقمًا أو نقطة أو شرطة.');
 if(!password)throw Error('أدخل كلمة المرور.');
 if(CONFIG.ALLOWED_ROLES.indexOf(role)<0)throw Error('الصلاحية غير صحيحة.');
 const sh=sheet_(USERS),v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++)if(String(v[i][2]).toLowerCase()===username)throw Error('اسم المستخدم مستخدم بالفعل.');
 const salt=Utilities.getUuid();sh.appendRow([Utilities.getUuid(),name,username,salt,hashPassword_(password,salt),role,true,new Date().toISOString()]);return {ok:true,message:'تم إنشاء المستخدم.'}
}
function listUsers_(u){requireRole_(u,['admin']);const v=sheet_(USERS).getDataRange().getValues(),users=[];for(let i=1;i<v.length;i++)if(v[i][0])users.push({id:String(v[i][0]),name:String(v[i][1]),username:String(v[i][2]),role:String(v[i][5]),active:truth_(v[i][6])});return {ok:true,users:users}}
function requireRole_(u,roles){if(!u||roles.indexOf(u.role)<0)throw Error('ليس لديك صلاحية لتنفيذ هذا الإجراء.')}
function hashPassword_(password,salt){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,salt+'|'+password,Utilities.Charset.UTF_8);return bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('')}
function truth_(v){return v===true||String(v).toLowerCase()==='true'}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.appendRow(headers);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#123b35').setFontColor('#ffffff');return sh}
function sheet_(name){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')||CONFIG.SPREADSHEET_ID;if(!id)throw Error('لم يتم إعداد قاعدة البيانات. شغّل setup() أولًا.');const sh=SpreadsheetApp.openById(id).getSheetByName(name);if(!sh)throw Error('الورقة غير موجودة: '+name);return sh}
function out_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}
function safeError_(e){const m=String(e&&e.message||e||'حدث خطأ غير معروف.');return m.length>240?m.slice(0,240):m}
