//==============================================================================
// 프로그램 설명 : 통합검진관리 작업전환 Function Unit
// 시스템        : 통합검진관리
// 수정일자      : 1999.07.13
// 수정자        : 강철호
// 수정내용      :
// 참고사항      :
//==============================================================================
unit MdmsCall;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs;

function  GF_MdmsCall(sProgId : String) : Boolean;

implementation

uses ZZMenu, ZZfunc,
     //입력 프로그램
     LD1I01,                                         LD1I07,                 LD1I10,
     LD1I11, LD1I12, LD1I13, LD1I14, LD1I15, LD1I16, LD1I17, LD1I18, LD1I19,
     LD2I01, LD2I02, LD2I03, LD2I04, LD2I05, LD2I06,
     LD3I06, LD3I07, LD3I08, LD3I09, LD3I10, LD3I11, LD3I12, ld3i14,
     LD4I01, LD4I02, LD4I03,
     LD5I01, LD5I02,
     LD6I01, LD6I02,
     LD9I01,
     LDAI01, LDAI02,

     //조회 프로그램
     LD2Q01, LD2Q02, LD2Q03, LD2Q04, LD2Q05,         LD2Q07, LD2Q08, LD2Q09, LD2Q10,
     LD2Q11, LD2Q12, LD2Q13, LD2Q14, LD2Q15, LD2Q16, LD2Q17, LD2Q18, LD2Q19,
     LD4Q01, LD4Q02, LD4Q03, LD4Q04,         LD4Q06,         LD4Q08, LD4Q09,
             LD4Q12, LD4Q13, LD4Q14, LD4Q15, LD4Q16, LD4Q17, LD4Q18, LD4Q19, LD4Q20,
     LD4Q21, LD4Q22, LD4Q23, LD4Q24, LD4Q25, LD4Q26, LD4Q27, LD4Q28, LD4Q29, LD4Q30,
     LD4Q31, LD4Q32, LD4Q33, LD4Q34, LD4Q35, LD4Q36, LD4Q37, LD4Q38, LD4Q39, LD4Q40,
     LD4Q41, LD4Q42, LD4Q43, LD4Q44, LD4Q45, LD4Q46, LD4Q47, LD4Q48, LD4Q49, LD4Q50,
     LD4Q51, LD4Q52, LD4Q53, LD4Q54, LD4Q55, LD4Q56, LD4Q57, LD4Q58, LD4Q59, LD4Q60,
     LD4Q61, LD4Q62, LD4Q63, LD4Q64, LD4Q65, LD4Q66, LD4Q67, LD4Q68, LD4Q69, LD4Q70,
     LD4Q71, LD4Q72, LD4Q73, LD4Q74, LD4Q75, LD4Q76, LD4Q77, LD4Q78, LD4Q79, LD4Q80,
     LD4Q81, LD4Q82, LD4Q83, LD4Q84, LD4Q85,
     LD5Q01, LD5Q02, LD5Q03, LD5Q04, LD5Q05, LD5Q06, LD5Q07, LD5Q08, LD5Q09, LD5Q10,
     LD5Q11, LD5Q12, LD5Q13,         LD5Q15, LD5Q16, LD5Q17,
     LD6Q01, LD6Q02, LD6Q03, LD6Q04,
     LD7Q01, LD7P01, LD7Q03, LD7P03, LD7Q04, LD7Q05, LD7Q06, LD7Q07,
     LD8Q01, LD8Q02, LD8Q03, LD8Q04, LD8Q05, LD8Q06, LD8Q07, LD8Q08, LD8Q09, LD8Q10,
     LD9Q01, LD9Q02, LD9Q03, LD9Q04, LD9Q05,

     //출력 프로그램
     LD8P01, LD8P02, LD8P03, LD8P04, LD5P14,
     LDAP01;
//==============================================================================
// 기능 : Program ID를 이용해서 해당 Form을 Create & Show한다.
//==============================================================================
function GF_MdmsCall(sProgId : String) : Boolean;
var sGubn_part : String;
begin
   Result := True;

   if (sProgId = 'LD1I01') or (sProgId = 'LD1I02') or
      (sProgId = 'LD1I03') or (sProgId = 'LD1I04') or
      (sProgId = 'LD1I05') or (sProgId = 'LD1I06') or
      (sProgId = 'LD1I08') or (sProgId = 'LD1I09') then
   begin
      Application.CreateForm(TfrmLD1I01, frmLD1I01);

      // Part구분을 전달
      if      sProgId = 'LD1I01' then sGubn_part := '02'    // 혈액학
      else if sProgId = 'LD1I02' then sGubn_part := '01'    // 생화학
      else if sProgId = 'LD1I03' then sGubn_part := '07'    // 혈청학
      else if sProgId = 'LD1I04' then sGubn_part := '03'    // URIN
      else if sProgId = 'LD1I05' then sGubn_part := '04'    // RIA
      else if sProgId = 'LD1I06' then sGubn_part := '05'    // EIA
      else if sProgId = 'LD1I08' then sGubn_part := '08'    // 기타1
      else if sProgId = 'LD1I09' then sGubn_part := '09';   // 기타2

      frmLD1I01.UP_EnvSet(sGubn_part);
      exit;
   end
   else if sProgId = 'LD1I07' then
   begin
      Application.CreateForm(TfrmLD1I07, frmLD1I07);
      exit;
   end
   else if sProgId = 'LD1I10' then
   begin
      Application.CreateForm(TfrmLD1I10, frmLD1I10);
      exit;
   end
   else if sProgId = 'LD1I11' then
   begin
      Application.CreateForm(TfrmLD1I11, frmLD1I11);
      exit;
   end
   else if sProgId = 'LD1I12' then
   begin
      Application.CreateForm(TfrmLD1I12, frmLD1I12);
      exit;
   end
   else if sProgId = 'LD1I13' then
   begin
      Application.CreateForm(TfrmLD1I13, frmLD1I13);
      exit;
   end
   else if sProgId = 'LD1I14' then
   begin
      Application.CreateForm(TfrmLD1I14, frmLD1I14);
      exit;
   end
   else if sProgId = 'LD1I15' then
   begin
      Application.CreateForm(TfrmLD1I15, frmLD1I15);
      exit;
   end
   else if sProgId = 'LD1I16' then
   begin
      Application.CreateForm(TfrmLD1I16, frmLD1I16);
      exit;
   end
   else if sProgId = 'LD1I17' then
   begin
      Application.CreateForm(TfrmLD1I17, frmLD1I17);
      exit;
   end
   else if sProgId = 'LD2I01' then
   begin
      Application.CreateForm(TfrmLD2I01, frmLD2I01);
      exit;
   end
   else if sProgId = 'LD2I02' then
   begin
      Application.CreateForm(TfrmLD2I02, frmLD2I02);
      exit;
   end
   else if sProgId = 'LD2I03' then
   begin
      Application.CreateForm(TfrmLD2I03, frmLD2I03);
      exit;
   end
   else if sProgId = 'LD2I04' then
   begin
      Application.CreateForm(TfrmLD2I04, frmLD2I04);
      exit;
   end
   else if sProgId = 'LD2I05' then
   begin
      Application.CreateForm(TfrmLD2I05, frmLD2I05);
      exit;
   end
   else if sProgId = 'LD2I06' then
   begin
      Application.CreateForm(TfrmLD2I06, frmLD2I06);
      exit;
   end
   else if sProgId = 'LD2Q01' then
   begin
      Application.CreateForm(TfrmLD2Q01, frmLD2Q01);
      exit;
   end
   else if sProgId = 'LD2Q02' then
   begin
      Application.CreateForm(TfrmLD2Q02, frmLD2Q02);
      exit;
   end
   else if sProgId = 'LD2Q03' then
   begin
      Application.CreateForm(TfrmLD2Q03, frmLD2Q03);
      exit;
   end
   else if sProgId = 'LD2Q04' then
   begin
      Application.CreateForm(TfrmLD2Q04, frmLD2Q04);
      exit;
   end
   else if sProgId = 'LD2Q05' then
   begin
      Application.CreateForm(TfrmLD2Q05, frmLD2Q05);
      exit;
   end

   else if sProgId = 'LD2Q07' then
   begin
      Application.CreateForm(TfrmLD2Q07, frmLD2Q07);
      exit;
   end
   else if sProgId = 'LD2Q08' then
   begin
      Application.CreateForm(TfrmLD2Q08, frmLD2Q08);
      exit;
   end
   else if sProgId = 'LD2Q09' then
   begin
      Application.CreateForm(TfrmLD2Q09, frmLD2Q09);
      exit;
   end
   else if sProgId = 'LD2Q10' then
   begin
      Application.CreateForm(TfrmLD2Q10, frmLD2Q10);
      exit;
   end
   else if sProgId = 'LD2Q11' then
   begin
      Application.CreateForm(TfrmLD2Q11, frmLD2Q11);
      exit;
   end
   else if sProgId = 'LD2Q12' then
   begin
      Application.CreateForm(TfrmLD2Q12, frmLD2Q12);
      exit;
   end
   else if sProgId = 'LD3I06' then
   begin
      Application.CreateForm(TfrmLD3I06, frmLD3I06);
      exit;
   end
   else if sProgId = 'LD3I07' then
   begin
      Application.CreateForm(TfrmLD3I07, frmLD3I07);
      exit;
   end
   else if sProgId = 'LD3I08' then
   begin
      Application.CreateForm(TfrmLD3I08, frmLD3I08);
      exit;
   end
   else if sProgId = 'LD3I09' then
   begin
      Application.CreateForm(TfrmLD3I09, frmLD3I09);
      exit;
   end
   else if sProgId = 'LD3I10' then
   begin
      Application.CreateForm(TfrmLD3I10, frmLD3I10);
      exit;
   end
   else if sProgId = 'LD4I01' then
   begin
      Application.CreateForm(TfrmLD4I01, frmLD4I01);
      exit;
   end
   else if sProgId = 'LD4I02' then
   begin
      Application.CreateForm(TfrmLD4I02, frmLD4I02);
      exit;
   end
   else if sProgId = 'LD4I03' then
   begin
      Application.CreateForm(TfrmLD4I03, frmLD4I03);
      exit;
   end
   else if sProgId = 'LD4Q01' then
   begin
      Application.CreateForm(TfrmLD4Q01, frmLD4Q01);
      exit;
   end
   else if sProgId = 'LD4Q02' then
   begin
      Application.CreateForm(TfrmLD4Q02, frmLD4Q02);
      exit;
   end
   else if sProgId = 'LD4Q03' then
   begin
      Application.CreateForm(TfrmLD4Q03, frmLD4Q03);
      exit;
   end

   else if sProgId = 'LD4Q04' then
   begin
      Application.CreateForm(TfrmLD4Q04, frmLD4Q04);
      exit;
   end
   else if sProgId = 'LD4Q06' then
   begin
      Application.CreateForm(TfrmLD4Q06, frmLD4Q06);
      exit;
   end
   else if sProgId = 'LD4Q08' then
   begin
      Application.CreateForm(TfrmLD4Q08, frmLD4Q08);
      exit;
   end
   else if sProgId = 'LD4Q09' then
   begin
      Application.CreateForm(TfrmLD4Q09, frmLD4Q09);
      exit;
   end
   else if sProgId = 'LD4Q12' then
   begin
      Application.CreateForm(TfrmLD4Q12, frmLD4Q12);
      exit;
   end
   else if sProgId = 'LD4Q13' then
   begin
      Application.CreateForm(TfrmLD4Q13, frmLD4Q13);
      exit;
   end
   else if sProgId = 'LD4Q14' then
   begin
      Application.CreateForm(TfrmLD4Q14, frmLD4Q14);
      exit;
   end
   else if sProgId = 'LD4Q15' then
   begin
      Application.CreateForm(TfrmLD4Q15, frmLD4Q15);
      exit;
   end


   else if sProgId = 'LD4Q19' then
   begin
      Application.CreateForm(TfrmLD4Q19, frmLD4Q19);
      exit;
   end
   else if sProgId = 'LD4Q20' then
   begin
      Application.CreateForm(TfrmLD4Q20, frmLD4Q20);
      exit;
   end
   else if sProgId = 'LD4Q21' then
   begin
      Application.CreateForm(TfrmLD4Q21, frmLD4Q21);
      exit;
   end
   else if sProgId = 'LD4Q22' then
   begin
      Application.CreateForm(TfrmLD4Q22, frmLD4Q22);
      exit;
   end
   else if sProgId = 'LD4Q23' then
   begin
      Application.CreateForm(TfrmLD4Q23, frmLD4Q23);
      exit;
   end
   else if sProgId = 'LD4Q24' then
   begin
      Application.CreateForm(TfrmLD4Q24, frmLD4Q24);
      exit;
   end
   else if sProgId = 'LD4Q25' then
   begin
      Application.CreateForm(TfrmLD4Q25, frmLD4Q25);
      exit;
   end
   else if sProgId = 'LD4Q26' then
   begin
      Application.CreateForm(TfrmLD4Q26, frmLD4Q26);
      exit;
   end
   else if sProgId = 'LD4Q27' then
   begin
      Application.CreateForm(TfrmLD4Q27, frmLD4Q27);
      exit;
   end
   else if sProgId = 'LD4Q28' then
   begin
      Application.CreateForm(TfrmLD4Q28, frmLD4Q28);
      exit;
   end
   else if sProgId = 'LD4Q29' then
   begin
      Application.CreateForm(TfrmLD4Q29, frmLD4Q29);
      exit;
   end
   else if sProgId = 'LD4Q30' then
   begin
      Application.CreateForm(TfrmLD4Q30, frmLD4Q30);
      exit;
   end
   else if sProgId = 'LD4Q31' then
   begin
      Application.CreateForm(TfrmLD4Q31, frmLD4Q31);
      exit;
   end
   else if sProgId = 'LD4Q32' then
   begin
      Application.CreateForm(TfrmLD4Q32, frmLD4Q32);
      exit;
   end
   else if sProgId = 'LD4Q33' then
   begin
      Application.CreateForm(TfrmLD4Q33, frmLD4Q33);
      exit;
   end
   else if sProgId = 'LD4Q36' then
   begin
      Application.CreateForm(TfrmLD4Q36, frmLD4Q36);
      exit;
   end
   else if sProgId = 'LD4Q37' then
   begin
      Application.CreateForm(TfrmLD4Q37, frmLD4Q37);
      exit;
   end
   else if sProgId = 'LD4Q38' then
   begin
      Application.CreateForm(TfrmLD4Q38, frmLD4Q38);
      exit;
   end
   else if sProgId = 'LD4Q39' then
   begin
      Application.CreateForm(TfrmLD4Q39, frmLD4Q39);
      exit;
   end
   else if sProgId = 'LD4Q40' then
   begin
      Application.CreateForm(TfrmLD4Q40, frmLD4Q40);
      exit;
   end
   else if sProgId = 'LD4Q41' then
   begin
      Application.CreateForm(TfrmLD4Q41, frmLD4Q41);
      exit;
   end
   else if sProgId = 'LD4Q43' then
   begin
      Application.CreateForm(TfrmLD4Q43, frmLD4Q43);
      exit;
   end
   else if sProgId = 'LD4Q44' then
   begin
      Application.CreateForm(TfrmLD4Q44, frmLD4Q44);
      exit;
   end
   else if sProgId = 'LD4Q45' then
   begin
      Application.CreateForm(TfrmLD4Q45, frmLD4Q45);
      exit;
   end
   else if sProgId = 'LD4Q46' then
   begin
      Application.CreateForm(TfrmLD4Q46, frmLD4Q46);
      exit;
   end
   else if sProgId = 'LD4Q47' then
   begin
      Application.CreateForm(TfrmLD4Q47, frmLD4Q47);
      exit;
   end
   else if sProgId = 'LD4Q48' then
   begin
      Application.CreateForm(TfrmLD4Q48, frmLD4Q48);
      exit;
   end
   else if sProgId = 'LD4Q49' then
   begin
      Application.CreateForm(TfrmLD4Q49, frmLD4Q49);
      exit;
   end
   else if sProgId = 'LD4Q50' then
   begin
      Application.CreateForm(TfrmLD4Q50, frmLD4Q50);
      exit;
   end
   else if sProgId = 'LD4Q51' then
   begin
      Application.CreateForm(TfrmLD4Q51, frmLD4Q51);
      exit;
   end
   else if sProgId = 'LD4Q52' then
   begin
      Application.CreateForm(TfrmLD4Q52, frmLD4Q52);
      exit;
   end
   else if sProgId = 'LD4Q53' then
   begin
      Application.CreateForm(TfrmLD4Q53, frmLD4Q53);
      exit;
   end
   else if sProgId = 'LD4Q54' then
   begin
      Application.CreateForm(TfrmLD4Q54, frmLD4Q54);
      exit;
   end
   else if sProgId = 'LD4Q55' then
   begin
      Application.CreateForm(TfrmLD4Q55, frmLD4Q55);
      exit;
   end
   else if sProgId = 'LD4Q56' then
   begin
      Application.CreateForm(TfrmLD4Q56, frmLD4Q56);
      exit;
   end
   else if sProgId = 'LD4Q57' then
   begin
      Application.CreateForm(TfrmLD4Q57, frmLD4Q57);
      exit;
   end
   else if sProgId = 'LD4Q58' then
   begin
      Application.CreateForm(TfrmLD4Q58, frmLD4Q58);
      exit;
   end
   else if sProgId = 'LD4Q59' then
   begin
      Application.CreateForm(TfrmLD4Q59, frmLD4Q59);
      exit;
   end
   else if sProgId = 'LD4Q60' then
   begin
      Application.CreateForm(TfrmLD4Q60, frmLD4Q60);
      exit;
   end
   else if sProgId = 'LD4Q61' then
   begin
      Application.CreateForm(TfrmLD4Q61, frmLD4Q61);
      exit;
   end
   else if sProgId = 'LD4Q62' then
   begin
      Application.CreateForm(TfrmLD4Q62, frmLD4Q62);
      exit;
   end
   else if sProgId = 'LD4Q63' then
   begin
      Application.CreateForm(TfrmLD4Q63, frmLD4Q63);
      exit;
   end
   else if sProgId = 'LD4Q64' then
   begin
      Application.CreateForm(TfrmLD4Q64, frmLD4Q64);
      exit;
   end
   else if sProgId = 'LD4Q65' then
   begin
      Application.CreateForm(TfrmLD4Q65, frmLD4Q65);
      exit;
   end
   else if sProgId = 'LD4Q66' then
   begin
      Application.CreateForm(TfrmLD4Q66, frmLD4Q66);
      exit;
   end
   else if sProgId = 'LD4Q67' then
   begin
      Application.CreateForm(TfrmLD4Q67, frmLD4Q67);
      exit;
   end
   else if sProgId = 'LD4Q68' then
   begin
      Application.CreateForm(TfrmLD4Q68, frmLD4Q68);
      exit;
   end
   else if sProgId = 'LD4Q69' then
   begin
      Application.CreateForm(TfrmLD4Q69, frmLD4Q69);
      exit;
   end
   else if sProgId = 'LD4Q70' then
   begin
      Application.CreateForm(TfrmLD4Q70, frmLD4Q70);
      exit;
   end
   else if sProgId = 'LD4Q71' then
   begin
      Application.CreateForm(TfrmLD4Q71, frmLD4Q71);
      exit;
   end
   else if sProgId = 'LD4Q72' then
   begin
      Application.CreateForm(TfrmLD4Q72, frmLD4Q72);
      exit;
   end
   else if sProgId = 'LD5I01' then
   begin
      Application.CreateForm(TfrmLD5I01, frmLD5I01);
      exit;
   end
   else if sProgId = 'LD5I02' then
   begin
      Application.CreateForm(TfrmLD5I02, frmLD5I02);
      exit;
   end
   else if sProgId = 'LD5Q01' then
   begin
      Application.CreateForm(TfrmLD5Q01, frmLD5Q01);
      exit;
   end
   else if sProgId = 'LD5Q02' then
   begin
      Application.CreateForm(TfrmLD5Q02, frmLD5Q02);
      exit;
   end
   else if sProgId = 'LD5Q03' then
   begin
      Application.CreateForm(TfrmLD5Q03, frmLD5Q03);
      exit;
   end
   else if sProgId = 'LD5Q04' then
   begin
      Application.CreateForm(TfrmLD5Q04, frmLD5Q04);
      exit;
   end
   else if sProgId = 'LD5Q05' then
   begin
      Application.CreateForm(TfrmLD5Q05, frmLD5Q05);
      exit;
   end
   else if sProgId = 'LD5Q06' then
   begin
      Application.CreateForm(TfrmLD5Q06, frmLD5Q06);
      exit;
   end
   else if sProgId = 'LD5Q07' then
   begin
      Application.CreateForm(TfrmLD5Q07, frmLD5Q07);
      exit;
   end
   else if sProgId = 'LD5Q08' then
   begin
      Application.CreateForm(TfrmLD5Q08, frmLD5Q08);
      exit;
   end
   else if sProgId = 'LD5Q09' then
   begin
      Application.CreateForm(TfrmLD5Q09, frmLD5Q09);
      exit;
   end
   else if sProgId = 'LD5Q10' then
   begin
      Application.CreateForm(TfrmLD5Q10, frmLD5Q10);
      exit;
   end
   else if sProgId = 'LD5Q11' then
   begin
      Application.CreateForm(TfrmLD5Q11, frmLD5Q11);
      exit;
   end
   else if sProgId = 'LD5Q12' then
   begin
      Application.CreateForm(TfrmLD5Q12, frmLD5Q12);
      exit;
   end
   else if sProgId = 'LD6I01' then
   begin
      Application.CreateForm(TfrmLD6I01, frmLD6I01);
      exit;
   end
   else if sProgId = 'LD6I02' then
   begin
      Application.CreateForm(TfrmLD6I02, frmLD6I02);
      exit;
   end
   else if sProgId = 'LD6Q01' then
   begin
      Application.CreateForm(TfrmLD6Q01, frmLD6Q01);
      exit;
   end
   else if sProgId = 'LD6Q02' then
   begin
      Application.CreateForm(TfrmLD6Q02, frmLD6Q02);
      exit;
   end
   else if sProgId = 'LD6Q03' then
   begin
      Application.CreateForm(TfrmLD6Q03, frmLD6Q03);
      exit;
   end
   else if sProgId = 'LD6Q04' then
   begin
      Application.CreateForm(TfrmLD6Q04, frmLD6Q04);
      exit;
   end
   else if sProgId = 'LD7Q01' then
   begin
      Application.CreateForm(TfrmLD7Q01, frmLD7Q01);
      exit;
   end
   else if sProgId = 'LD7P01' then
   begin
      frmLD7P01 := TfrmLD7P01.Create(frmMenu);
      frmLD7P01.ShowModal;
      exit;
   end
   else if sProgId = 'LD7Q03' then
   begin
      Application.CreateForm(TfrmLD7Q03, frmLD7Q03);
      exit;
   end
   else if sProgId = 'LD7Q04' then
   begin
      Application.CreateForm(TfrmLD7Q04, frmLD7Q04);
      exit;
   end
   else if sProgId = 'LD7P03' then
   begin
      frmLD7P03 := TfrmLD7P03.Create(frmMenu);
      frmLD7P03.ShowModal;
      exit;
   end
   else if sProgId = 'LD8Q01' then
   begin
      Application.CreateForm(TfrmLD8Q01, frmLD8Q01);
      exit;
   end
   else if sProgId = 'LD8Q02' then
   begin
      Application.CreateForm(TfrmLD8Q02, frmLD8Q02);
      exit;
   end
   else if sProgId = 'LD8Q03' then
   begin
      Application.CreateForm(TfrmLD8Q03, frmLD8Q03);
      exit;
   end
   else if sProgId = 'LD8Q04' then
   begin
      Application.CreateForm(TfrmLD8Q04, frmLD8Q04);
      exit;
   end
   else if sProgId = 'LD8Q05' then
   begin
      Application.CreateForm(TfrmLD8Q05, frmLD8Q05);
      exit;
   end
   else if sProgId = 'LD8Q06' then
   begin
      Application.CreateForm(TfrmLD8Q06, frmLD8Q06);
      exit;
   end
   else if sProgId = 'LD8Q07' then
   begin
      Application.CreateForm(TfrmLD8Q07, frmLD8Q07);
      exit;
   end
   else if sProgId = 'LD8Q08' then
   begin
      Application.CreateForm(TfrmLD8Q08, frmLD8Q08);
      exit;
   end
   else if sProgId = 'LD8Q09' then
   begin
      Application.CreateForm(TfrmLD8Q09, frmLD8Q09);
      exit;
   end
   else if sProgId = 'LD8P01' then
   begin
      frmLD8P01 := TfrmLD8P01.Create(frmMenu);
      frmLD8P01.ShowModal;
      exit;
   end
   else if sProgId = 'LD8P02' then
   begin
      frmLD8P02 := TfrmLD8P02.Create(frmMenu);
      frmLD8P02.ShowModal;
      exit;
   end
   else if sProgId = 'LD3I11' then
   begin
      frmLD3I11 := TfrmLD3I11.Create(frmMenu);
      exit;
   end
   else if sProgId = 'LD2Q14' then
   begin
      frmLD2Q14 := TfrmLD2Q14.Create(frmMenu);
      exit;
   end
   else if sProgId = 'LD2Q15' then
   begin
      frmLD2Q15 := TfrmLD2Q15.Create(frmMenu);
      exit;
   end
   else if sProgId = 'LD8P03' then
   begin
      frmLD8P03 := TfrmLD8P03.Create(frmMenu);
      frmLD8P03.ShowModal;
      exit;
   end
   else if sProgId = 'LD8P04' then
   begin
      frmLD8P04 := TfrmLD8P04.Create(frmMenu);
      frmLD8P04.ShowModal;
      exit;
   end
   else if sProgId = 'LD9I01' then
   begin
      Application.CreateForm(TfrmLD9I01, frmLD9I01);
      exit;
   end
   else if sProgId = 'LD9Q01' then
   begin
      Application.CreateForm(TfrmLD9Q01, frmLD9Q01);
      exit;
   end
   else if sProgId = 'LD9Q02' then
   begin
      Application.CreateForm(TfrmLD9Q02, frmLD9Q02);
      exit;
   end
   else if sProgId = 'LD9Q03' then
   begin
      Application.CreateForm(TfrmLD9Q03, frmLD9Q03);
      exit;
   end
   else if sProgId = 'LD9Q04' then
   begin
      Application.CreateForm(TfrmLD9Q04, frmLD9Q04);
      exit;
   end
   else if sProgId = 'LD4Q73' then
   begin
      Application.CreateForm(TfrmLD4Q73, frmLD4Q73);
      exit;
   end
   else if sProgId = 'LD4Q74' then
   begin
      Application.CreateForm(TfrmLD4Q74, frmLD4Q74);
      exit;
   end
   else if sProgId = 'LD4Q79' then
   begin
      Application.CreateForm(TfrmLD4Q79, frmLD4Q79);
      exit;
   end
   else if sProgId = 'LD5Q13' then
   begin
      Application.CreateForm(TfrmLD5Q13, frmLD5Q13);
      exit;
   end
   else if sProgId = 'LD5Q15' then
   begin
      Application.CreateForm(TfrmLD5Q15, frmLD5Q15);
      exit;
   end
   else if sProgId = 'LD5P14' then
   begin
      Application.CreateForm(TfrmLD5P14, frmLD5P14); 
      frmLD5P14.ShowModal;
      exit;
   end
   else if sProgId = 'LD2Q13' then
   begin
      Application.CreateForm(TfrmLD2Q13, frmLD2Q13);
      exit;
   end
   else if sProgId = 'LD4Q80' then
   begin
      Application.CreateForm(TfrmLD4Q80, frmLD4Q80);
      exit;
   end
   else if sProgId = 'LD4Q75' then
   begin
      Application.CreateForm(TfrmLD4Q75, frmLD4Q75);
      exit;
   end
   else if sProgId = 'LD4Q76' then
   begin
      Application.CreateForm(TfrmLD4Q76, frmLD4Q76);
      exit;
   end
   else if sProgId = 'LD4Q42' then
   begin
      Application.CreateForm(TfrmLD4Q42, frmLD4Q42);
      exit;
   end
   else if sProgId = 'LD3I12' then
   begin
      Application.CreateForm(TfrmLD3I12, frmLD3I12);
      exit;
   end
   else if sProgId = 'LD9Q05' then
   begin
      Application.CreateForm(TfrmLD9Q05, frmLD9Q05);
      exit;
   end
   else if sProgId = 'LD4Q16' then
   begin
      Application.CreateForm(TfrmLD4Q16, frmLD4Q16);
      exit;
   end
   else if sProgId = 'LD1I18' then
   begin
      Application.CreateForm(TfrmLD1I18, frmLD1I18);
      exit;
   end
   else if sProgId = 'LD4Q17' then
   begin
      Application.CreateForm(TfrmLD4Q17, frmLD4Q17);
      exit;
   end
   else if sProgId = 'LD4Q18' then
   begin
      Application.CreateForm(TfrmLD4Q18, frmLD4Q18);
      exit;
   end
   else if sProgId = 'LD2Q17' then
   begin
      Application.CreateForm(TfrmLD2Q17, frmLD2Q17);
      exit;
   end
   else if sProgId = 'LDAI01' then
   begin
      Application.CreateForm(TfrmLDAI01, frmLDAI01);
      exit;
   end
   else if sProgId = 'LDAI02' then
   begin
      Application.CreateForm(TfrmLDAI02, frmLDAI02);
      exit;
   end
   else if sProgId = 'LD5Q16' then
   begin
      Application.CreateForm(TfrmLD5Q16, frmLD5Q16);
      exit;
   end
   else if sProgId = 'LD4Q77' then
   begin
      Application.CreateForm(TfrmLD4Q77, frmLD4Q77);
      exit;
   end
   else if sProgId = 'LD4Q81' then
   begin
      Application.CreateForm(TfrmLD4Q81, frmLD4Q81);
      exit;
   end
   else if sProgId = 'LD2Q16' then
   begin
      Application.CreateForm(TfrmLD2Q16, frmLD2Q16);
      exit;
   end
   else if sProgId = 'LD7Q05' then
   begin
      Application.CreateForm(TfrmLD7Q05, frmLD7Q05);
      exit;
   end
   else if sProgId = 'LD4Q82' then
   begin
      Application.CreateForm(TfrmLD4Q82, frmLD4Q82);
      exit;
   end
   else if sProgId = 'LD4Q83' then
   begin
      Application.CreateForm(TfrmLD4Q83, frmLD4Q83);
      exit;
   end
   else if sProgId = 'LDAP01' then
   begin
      Application.CreateForm(TfrmLDAP01, frmLDAP01);
      exit;
   end
   else if sProgId = 'LD8Q10' then
   begin
      Application.CreateForm(TfrmLD8Q10, frmLD8Q10);
      exit;
   end
   else if sProgId = 'LD1I19' then
   begin
      Application.CreateForm(TfrmLD1I19, frmLD1I19);
      exit;
   end
   else if sProgId = 'LD2Q18' then
   begin
      Application.CreateForm(TfrmLD2Q18, frmLD2Q18);
      exit;
   end
   else if sProgId = 'LD4Q34' then
   begin
      Application.CreateForm(TfrmLD4Q34, frmLD4Q34);
      exit;
   end
   else if sProgId = 'LD2Q19' then
   begin
      Application.CreateForm(TfrmLD2Q19, frmLD2Q19);
      exit;
   end
   else if sProgId = 'LD7Q06' then
   begin
      Application.CreateForm(TfrmLD7Q06, frmLD7Q06);
      exit;
   end
   else if sProgId = 'LD7Q07' then
   begin
      Application.CreateForm(TfrmLD7Q07, frmLD7Q07);
      exit;
   end
   else if sProgId = 'LD4Q35' then
   begin
      Application.CreateForm(TfrmLD4Q35, frmLD4Q35);
      exit;
   end
   else if sProgId = 'LD5Q17' then
   begin
      Application.CreateForm(TfrmLD5Q17, frmLD5Q17);
      exit;
   end
   else if sProgId = 'LD3I14' then
   begin
      Application.CreateForm(TfrmLD3I14, frmLD3I14);
      exit;
   end
   else if sProgId = 'LD4Q84' then
   begin
      Application.CreateForm(TfrmLD4Q84, frmLD4Q84);
      exit;
   end
   else if sProgId = 'LD4Q78' then
   begin
      Application.CreateForm(TfrmLD4Q78, frmLD4Q78);
      exit;
   end
   else if sProgId = 'LD4Q85' then
   begin
      Application.CreateForm(TfrmLD4Q85, frmLD4Q85);
      exit;
   end;
end;
end.
