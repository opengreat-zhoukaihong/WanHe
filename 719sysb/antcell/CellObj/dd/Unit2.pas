unit Unit2;

interface

uses
   Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,stdctrls;
type
    TPiont=record
     X:Double;
     Y:Double;
    end;
    spptr=^pp;
     pp=Record
      x,y:Double;
      link: spptr;
      q:integer;
     end;

var
   A,B,C,Z,O,H:TPiont;

      nangles:array of string;
      npionts:array of pp;
      kk,Sum,bz,bnm:integer;
      rec,strpoint:string;
      left,right,top,bottom,Max:Double;
procedure initdata ;
procedure initanglesinf;
procedure maincalc ;
procedure getstoreinf(angleinf:string;no:integer);
procedure bbianjie(var H:Tpiont; Z,A,B:Tpiont);
procedure AddPoint(A,Z,H:TPiont;ang:Double);
function  firstpoint(kkk:integer):integer;
function  ZhongxinPiont(A,B,C:TPiont):TPiont;
function  NNextPiont(A,B:TPiont;var H,Z:Tpiont):Integer;
function  NextPiont(A,B,Z:TPiont;var H:Tpiont):Integer;
function  direction(A,H,Z:Tpiont):integer;
function  bianjie( Var H:Tpiont;Z:Tpiont):integer;
procedure storedrawpiont(gg:string;no:integer);
procedure store_order(gg:string;no:integer);
procedure updatatbl(mm:string);

implementation
     uses Unit1;


function ZhongxinPiont(A,B,C:TPiont):TPiont;
var  k1,k2,x1,x2,y1,y2:Double;
     k:integer;
  begin
       K:=0;
       if b.y<>a.y then
       begin
         k1:=-(b.x-a.x)/(b.y-a.y);
         K:=k+1;
       end;

       if b.y<>c.y then
       begin
         k2:=-(b.x-c.x)/(b.y-c.y);
         K:=K+1;
       end;

       x1:=(a.x+b.x)/2;y1:=(a.y+b.y)/2;
       x2:=(c.x+b.x)/2;y2:=(c.y+b.y)/2;

       if  (k=1)then
       begin
        if b.y=a.y then
         begin
         z.x:=x1;  z.y:=k2*(z.x-x2)+y2;
         end ;
        if  b.y=c.y then
         begin
         z.x:=x2; z.y:=k1*(z.x-x1)+y1;
         end;
        end
        else

        if (k1<>k2) and (k=2)then
         begin
         z.x:=((k1*x1-k2*x2)+ (y2-y1))/(k1-k2);
         z.y:=k1*(z.x-x1)+y1;
         end;

       result:=z;
  end;

procedure AddPoint(A,Z,H:TPiont;ang:Double);
var  k1,k2,x1,y1:Double;
     k,x,y,ff:integer;
     str:string;
  begin

       K:=0;
       if (ang<>90)and (ang<>270) then
       begin
         k1:=sin(ang/180*3.1415926)/cos(ang/180*3.1415926);
         K:=k+1;
       end;

       if z.x<>h.x then
       begin
         k2:=(z.y-h.y)/(z.x-h.x);
         K:=K+1;
       end;

       ff:=0;
       if ((k1<>k2)and (k=2)) or (k=1)then
        begin
         if k=1 then
         begin

           if (ang=90)or (ang=270) then
           begin
             x1:=a.x;
             y1:=k2*(x1-z.x)+z.y;
             ff:=1;
           end
           else
           if  z.x=h.x then
           begin
             x1:=z.x;
             y1:=k1*(x1-a.x)+a.y;
             ff:=1;
           end;
         end
         else
         begin
           if abs(k1-k2)>0.0001 then
           begin
            x1:=((k1*a.x-k2*z.x)+ (z.y-a.y))/(k1-k2);
            y1:=k2*(x1-z.x)+z.y;
            ff:=1;
           end;
          end;

         if (ff=1) then
         if ((x1-h.x)*(x1-z.x)<0) or (((x1-h.x)*(x1-z.x)=0)and ((y1-h.y)*(y1-z.y)<0)) then
         begin
           if ((ang<180) and  (ang>0)  and ((y1-a.y)>0)) then  ff:=0;
           if ((ang>180) and (ang<360) and ((y1-a.y)<0)) then  ff:=0;
           if ((ang=0)   and ((x1-a.x)>0)) then  ff:=0;
           if ((ang=180) and ((x1-a.x)<0)) then  ff:=0;
           if (ff=0) then
           begin
            rec:=rec+floattostr(ang)+';';
            kk:=kk+1;

            x:= round(((x1-40)/1000+left1)*1000000);
            y:= round(((y1-40)/1000+bottom1)*1000000);
            strpoint:=strpoint+floattostr(x/1000000)+','+ floattostr(y/1000000)+',';
            x:= round(((a.x-40)/1000+left1)*1000000);
            y:= round(((a.y-40)/1000+bottom1)*1000000);
            strpoint:=strpoint+floattostr(x/1000000)+','+ floattostr(y/1000000)+','+'/';
            strpoint:=strpoint+floattostr(x/1000000)+','+ floattostr(y/1000000)+',';
            x:= round(((x1-40)/1000+left1)*1000000);
            y:= round(((y1-40)/1000+bottom1)*1000000);
            strpoint:=strpoint+floattostr(x/1000000)+','+ floattostr(y/1000000)+',';
           end;
         end;
     end;
  end;


procedure initdata ;
  var i:integer;stt:string;
      x1,y1,x,y:double;
  begin

    SetLength(npionts, 200);SetLength(nangles, 200);

    try
    with fmCellObj.quCellObj do
    begin
      close;
      ////////
      //DatabaseName := wExePath;

      /////////
      sql.clear;
      //sql.Add('select  LON,LAT,BEARING from pointtbl order by LON,LAT,BEARING');
      sql.Add('select  LON,LAT,BEARING from CellObj order by LON,LAT,BEARING');

      open;
      sum:=recordcount;
      SetLength(npionts, sum);SetLength(nangles, sum);

      first; i:=0;
      x1:=-1;y1:=-1; stt:='';
      while not EOF do
      begin
        x:=FieldByName('LON').ASFloat;
        y:=FieldByName('LAT').ASFloat;

       if (x1<>x) or (y1<>y) then
        begin
           npionts[i].x:=(x-left1)*1000+40;
           npionts[i].Y:=(y-bottom1)*1000+40;
           npionts[i].link:=nil;
           npionts[i].q:=1;
           x1:=x;y1:=y;
           if i>0 then  nangles[i-1]:=stt;
           stt:='';
           i:=i+1;
        end;
        stt:=stt+inttostr(FieldByName('BEARING').ASinteger)+',';
        next;
      end;
      nangles[i]:=stt;
      close;
    end;
    except
    end;
      sum:=i;
      SetLength(npionts, sum);SetLength(nangles, sum);
      left:=0;
      right:=(right1-left1)*1000+80;
      bottom:=0;
      top:=(top1-bottom1)*1000+80;
      MAX:=(left-right)*(left-right) +(top-bottom)*(top-bottom)
  end;


function firstpoint(kkk:integer):integer;
  Var  i,j,m:integer;
       dm,tdm:double;
  begin
      j:=kkk;
      dm:=MAx;
      for i:=0 to sum-1  do
      begin
        if  j=i  then continue;
        tdm:=(npionts[i].x-npionts[j].x)*(npionts[i].x-npionts[j].x)+
             (npionts[i].y-npionts[j].y)*(npionts[i].y-npionts[j].y);
        if dm>tdm then
         begin
         dm:=tdm;
         m:=i;
         end;
      end;
      result:=m;
  end;

procedure maincalc ;
  var  i,j,x,y,jj,cs,vn,vm,mm : integer;
       Tw,tq,ptr,tt:spptr;
       angles,ags,fx,fy,all : string;
  begin

   try
   initdata ;
   except
   end;

   initanglesinf;
   all:='';
   ags:='';

   for i:=0 to sum-1 do
      begin
      try

        for j:=0 to sum-1 do   npionts[j].q:=1;

        j:=firstpoint(i);
        npionts[j].q:=0;

        A.x:=npionts[i].x; A.y:=npionts[i].y;
        B.x:=npionts[j].x; B.y:=npionts[j].y;

        try
        jj:=NNextPiont(A,B,H,Z);
        except
        break;
        end;
        if KK=-2 then continue;

        try
        bz:=direction(A,H,Z);
        except
        break;
        end;

        C.x:=npionts[jj].x; C.y:=npionts[jj].y;

        new(ptr);
        npionts[i].link:=ptr;  ptr.q:=kk;
        ptr.link:=nil; ptr.x:=Z.x; ptr.y:=Z.y;  tt:=ptr;

        new(ptr);
        tt.link:=ptr;    ptr.q:=jj;
        ptr.link:=nil; ptr.x:=H.x; ptr.y:=H.y;  tt:=ptr;

        if kk=-1 then cs:=1 else cs:=0 ;

        repeat
           Z.x:=H.X;Z.y:=H.Y;

           if (kk=-1) and  (cs=0) then
           begin
             cs:=1;
             C.x:=npionts[jj].x;C.y:=npionts[jj].y;
             npionts[jj].q:=0;
             tw:=npionts[i].link;Tq:=tw.link;tw.link:=nil;tt:=TW;

             while Tq<>nil do
             begin
                npionts[i].link:=tq;
                tq:=npionts[i].link.link;
                npionts[i].link.link:=tw;
                tw:=npionts[i].link;
             end;

             B.x:=npionts[j].x;B.y:=npionts[j].y;
             Z.x:=ZhongxinPiont(A,B,C).x;Z.y:=ZhongxinPiont(A,B,C).y;

             if bz=1 then  bz:=0 else bz:=1;
           end;

           B.x:=C.X;B.y:=C.Y;
           try
           jj:=NextPiont(A,B,Z,H);
           except
           break;
           end;

           new(ptr);
           tt.link:=ptr;    ptr.q:=jj;
           ptr.link:=nil; ptr.x:=H.x; ptr.y:=H.y;  tt:=ptr;
           //Dispose(ptr);
           if (abs(H.x-npionts[i].link.x)<0.00001)  and  (abs(H.y-npionts[i].link.y)<0.00001)  then jj:=kk;
           C.x:=npionts[jj].x; C.y:=npionts[jj].y;

        until jj=kk ;  //repeat...unitl....

        if (kk=-1) and (tt.x<>npionts[i].link.x) and  (tt.y<>npionts[i].link.y)  then
        begin
           new(ptr);
           if  (TT.x=left) or  (TT.x=right) then
           begin
            ptr.x:=TT.x; ptr.y:=npionts[i].link.y;
           end;
           if (TT.y=bottom) or (TT.y=top)   then
           begin
            ptr.y:=TT.y; ptr.x:=npionts[i].link.x;
           end;
           tt.link:=ptr;
           ptr.link:=nil;  tt:=ptr;
          // Dispose(ptr);
        end;

        kk:=0;
        angles:=copy(nangles[i],3,length(nangles[i]));
        if (angles<>'') then jj:=strtoint(copy(nangles[i],1,1)) else jj:=0;
        getstoreinf(angles,i);

        if kk = jj then
          store_order(angles,i)
        else
        begin
          if jj=0  then strpoint:='/';
          if jj=2  then strpoint:='//';
          if jj=3  then strpoint:='///';
        end;

      except
      end;
      all := all + strpoint;

      end;

      updatatbl(all);

  end ;

procedure updatatbl(mm:string);
var  i,ll:integer;
     temp,tp:string;
begin
      tp := mm;
      ll := length(tp);
      //with main.Table1 do
      with fmCellObj.taCellObj do
      begin
        open;
        first;
        while not EOF do
        begin
          edit;
          temp := copy(tp,1,pos('/',tp)-1);
          i:=1;
          while temp<>'' do
          begin
            FieldbyName('LON_'+inttostr(i)).Asstring :=
              copy(temp,1,pos(',',temp)-1);
            temp := trim(copy(temp,pos(',',temp)+1,length(temp)));
            FieldbyName('LAT_'+inttostr(i)).Asstring :=
              copy(temp,1,pos(',',temp)-1);
            temp := trim(copy(temp,pos(',',temp)+1,length(temp)));
            i:=i+1;
          end;
          tp:=copy(tp,pos('/',tp)+1,ll);
          post();
          next;
        end;
        close;
      end;
end;

procedure initanglesinf;
var   j,i,a1,a,aa:integer;
      str,str1,str2:string;
  begin
    j:=-1;
    while true do
    begin
        j:=j+1;
        if j=sum then break;
        str:=nangles[j]; i:=1;
        if str='' then  continue;
        str2:=copy(str,1,pos(',',str)-1);
        a1:=strtoint(str2);
        aa:=a1;
        str2:='('+str2+')';
        str:=copy(str,pos(',',str)+1,length(str));
        while str<>'' do
        begin
         a:=strtoint(copy(str,1,pos(',',str)-1));
         if 90-(a+a1)/2<0 then a1:=a1-720;
         str2 := str2 + floattostr(90-(a+a1)/2) + '('+inttostr(a)+')';
         a1:=a;
         i:=i+1;
         str:=copy(str,pos(',',str)+1,length(str));
        end;
        if i>1 then
         begin
         if (360-a1)<aa then a:=2*aa-((360-a1)+aa)
                        else a:=2*a1+((360-a1)+aa);
         if 90-a/2<0 then a:=a-720;
         str2 := floattostr(90-a/2)+str2;
         end;
         if i=1 then nangles[j]:='' else nangles[j]:=inttostr(i)+'-'+str2;
    end;

end;

procedure getstoreinf(angleinf:string;no:integer);
var ags:string;
    mm:Double;
    i,x,y,x1,y1:integer;
    Tw,tq,ptr,tt:spptr;
begin
        i:=no;

        rec:='';
        strpoint:='';
        tw:=npionts[i].link;Tq:=tw.link;
        a.x:=npionts[i].x; a.y:=npionts[i].y;

        z.x:=tw.x ; z.y:=tw.y ;

        x1:= round(((z.x-40)/1000+left1)*1000000);
        y1:= round(((z.y-40)/1000+bottom1)*1000000);
        strpoint:=floattostr(x1/1000000)+','+ floattostr(y1/1000000)+',';
        while Tq<>nil do
        begin
          H.x:=tq.x ; H.y:=tq.y ;
          ags:=angleinf;
          while ags<>'' do
          begin
            mm:=strtofloat(copy(ags,1,pos('(',ags)-1));
            AddPoint(A,Z,H,mm);
            ags:=trim(copy(ags,pos(')',ags)+1,length(ags)));
          end;

          z.x:=h.x ; z.y:=h.y ;

          x1:= round(((z.x-40)/1000+left1)*1000000);
          y1:= round(((z.y-40)/1000+bottom1)*1000000);
          strpoint:=strpoint+floattostr(x1/1000000)+','+ floattostr(y1/1000000)+',';
          tw:=Tq; Tq:=tw.link;
        end;

        H.x:=z.x ; h.y:=z.y ;
        Z.x:=npionts[i].link.x; Z.y:=npionts[i].link.y ;

        ags:=angleinf;
        while ags<>'' do
        begin
            mm:=strtofloat(copy(ags,1,pos('(',ags)-1));
            AddPoint(A,Z,H,mm);
            ags:=trim(copy(ags,pos(')',ags)+1,length(ags)));
        end;

        if kk=0 then
        begin
        z.x:=npionts[i].link.x ; z.y:=npionts[i].link.y ;
        x1:= round(((z.x-40)/1000+left1)*1000000);
        y1:= round(((z.y-40)/1000+bottom1)*1000000);
        strpoint:=strpoint+floattostr(x1/1000000)+','+ floattostr(y1/1000000)+',';
        end;

end;

procedure store_order(gg:string;no:integer);
var w,st1,st,temp1,temp2:string;
    i:integer;
    j:integer;
begin
   if trim(gg)<>'' then
   begin
    w:=gg;temp1:='';temp2:='';
    st1:=copy(strpoint,1,pos('/',strpoint));
    strpoint:=copy(strpoint,pos('/',strpoint)+1,length(strpoint))+st1;
    if bz=0 then
    begin
      i:=0;
      while (rec<>'') and  (strpoint<>'') do
      begin
      st:=copy(strpoint,1,pos('/',strpoint));
      temp1:=st+temp1;
      strpoint:=trim(copy(strpoint,pos('/',strpoint)+1,length(strpoint)));

      st:=copy(rec,1,pos(';',rec));
      temp2:=st+temp2;
      rec:=trim(copy(rec,pos(';',rec)+1,length(rec)));
      i:=i+1 ;
      if i>4 then  break;
      end;
      strpoint:=temp1; rec:=temp2;
       if i>4 then showmessage( 'strpoint='+strpoint);
    end;

    st:=trim(copy(w,1,pos('(',w)-1));
    temp1:=trim(copy(rec,1,pos(';',rec)-1));
    i:=0;
    while st<>temp1 do
    begin
      temp1:=copy(strpoint,1,pos('/',strpoint));
      strpoint:=copy(strpoint,pos('/',strpoint)+1,length(strpoint))+temp1;
      temp1:=trim(copy(rec,1,pos(';',rec)-1));
      rec:=copy(rec,pos(';',rec)+1,length(rec))+temp1+';';
      i:=i+1;
      if i>4 then break;
    end;
    if i>4 then showmessage( 'strpoint='+strpoint);
   end
   else  strpoint:=strpoint+'/';
end;

procedure storedrawpiont(gg:string; no:integer);
var
   i,j,w2,w3 : integer;
    st1,st2,st3,w1,w,g,v0,vv,v : string;
begin;
   w:=gg;
   if bz=0 then
   begin
    w:='';
    while gg <>'' do
    begin
      w1:=copy(gg,1,pos(')',gg));
      w:=copy(w1,pos('(',w1),pos(')',w1))+copy(w1,1,pos('(',w1)-1)+w;
      gg:=trim(copy(gg,pos(')',gg)+1,length(gg)));
    end;
    w1:=copy(w,1,pos(')',w));
    w:=trim(copy(w,pos(')',w)+1,length(w)))+w1;
   end;
   V0:=copy(rec,1,pos(';',rec)-1);
   vv:=trim(copy(rec,1,pos('(',w)-1));
   v:='';
   i:=0;
   while (w<>'') and (V0<>vv) do
    begin
      i:=i+1;
      vv:=trim(copy(rec,1,pos('(',w)-1));
      V:=copy(w,1,pos(')',w))+v;
      W:=trim(copy(w,pos(')',w)+1,length(w)));
    end;
   w:=w+v;
   // showmessage('w='+w+char(13)+rec);
   st1:=copy(strpoint,1,pos('/',strpoint));
   strpoint:=copy(strpoint,pos('/',strpoint)+1,length(strpoint))+st1;

   st1:='';
   i:=0; w1:=inttostr(No);
   while (w<>'')  do
    begin
      V:=copy(w,1,pos(')',w));
      V:=copy(V,pos('(',v)+1,pos(')',v)-pos('(',v)-1);
      vv:=v;
      v:=v+'--'+copy(strpoint,1,pos('/',strpoint)-1);
      strpoint:=copy(strpoint,pos('/',strpoint)+1,length(strpoint));
      w:=trim(copy(w,pos(')',w)+1,length(w)));
      st1:=st1+V+char(13);
      i:=i+1;
      if i=5 then  break;
    end;
end;



function NNextPiont(A,B:TPiont;var H,Z:Tpiont):Integer;
  var   i,flag,gh:integer;
        LL,LL1,X1,Y1,k1:double;
  begin

      O.x:=(A.X+B.X)/2; O.y:=(A.y+B.y)/2;
      try
      if A.x<>B.x then K1:=(A.y-B.Y)/(A.x-b.X);
      except
        showmessage('error1');
      end;

      gh:=0; LL:=MAX;
      for i:=0 to Sum-1 do
      begin
         C.x:=npionts[i].x;
         C.y:=npionts[i].y;

         if ((C.x=A.x)and(C.y=A.y))OR((C.x=B.x)and(C.y=B.y))then continue;
         if (A.x=B.x)and(B.x=c.x) then continue;
         try
         if (C.x<>B.x) and (A.x<>B.x) and  (K1=(C.y-B.Y)/(C.x-b.X)) then continue;
         except
         showmessage('error');
         end;

         try
           H:=ZhongxinPiont(A,B,C);
         except
           showmessage('error3');
         end;

         LL1:=(H.x-O.x)*(H.x-O.x)+(H.y-O.y)*(H.y-O.y);
         if LL>LL1 then
         begin
          LL:=LL1;
          flag:=i;
         end;

      end;

      npionts[flag].q:=0;
      result:=flag;
      C.x:=npionts[flag].x;c.y:=npionts[flag].y;
      H := ZhongxinPiont(A,B,C);
      Z.x := O.X; Z.y:=O.y;
      try
        if bianjie(H,Z)=0 then gh:=1;
        except
        showmessage('error6');
      end;
      flag:=-1; LL:=MAX;
      for i:=0 to Sum-1 do
      begin
        C.x:=npionts[i].x;
        c.y := npionts[i].y;

        if ((C.x=A.x)and(C.y=A.y))OR((C.x=B.x)and(C.y=B.y))or (npionts[i].q=0)then continue;
        if (A.x=B.x)and(B.x=c.x) then continue;
        try
        if (C.x<>B.x) and (A.x<>B.x) and  (K1=(C.y-B.Y)/(C.x-b.X)) then continue;
        except
        showmessage('error4');
      end;
        Z:=ZhongxinPiont(A,B,C);
        if ((O.x-H.X)*(O.x-Z.X)<0) or ((O.y-H.Y)*(O.Y-Z.Y)<0) then
        begin
           LL1:=(Z.x-O.x)*(Z.x-O.x)+(Z.y-O.y)*(Z.y-O.y);
           if LL>LL1 then
            begin
            LL:=LL1;
            flag:=i;
            end;
        end;

      end;

      if flag = -1 then
      begin
         X1:=H.x;
         Y1:=H.y;
         H.x:=O.x;
         H.y:=O.y;
         Z.X:=x1;
         Z.Y:=y1;
         try
         Bz:=direction(A,H,Z);
         bbianjie(H,Z,A,B);
          except
        showmessage('error11');
        end;
         Z.x:=H.x;Z.y:=H.y;
         If gh=1 then kk:=-2  else   KK:=-1;
         H.x:=X1;H.y:=Y1;
      end
      else
      begin
         kk:=flag;
         C.x:=npionts[flag].x;c.y:=npionts[flag].y;
         Z:=ZhongxinPiont(A,B,C);
       //  showmessage('floa+++++++='+floattostr(Z.x)+'floa+++++++='+floattostr(Z.y));
         if (Z.x<left) or (Z.x>right) or (Z.Y>top) or (Z.Y<bottom) then
         begin
            X1:=H.x;Y1:=H.y; H.x:=Z.x;H.y:=Z.y; Z.x:=X1;Z.y:=y1;
            bianjie(H,Z);
            Z.x:=H.x;
            Z.y:=H.Y;
            H.x:=X1;
            H.y:=Y1;
            If gh=1 then
              kk := -2
            else
              KK:=-1;
         end
         else
         begin
         if (gh=1) then
         begin
            X1:=Z.x;Y1:=Z.y; Z.x:=H.x;Z.y:=H.Y; H.x:=X1;H.y:=Y1;
            npionts[result].q:=1;
            result:=kk;
            kk:=-1;
         end;
         end;
      end;
     // showmessage('floa+++++++='+floattostr(Z.x)+'floa+++++++='+floattostr(Z.y));
  end;


function NextPiont(A,B,Z:TPiont;var H:Tpiont):Integer;
  var  i,flag1,flag2,nm:integer;
       LL,LL2,LL1,k1:double;
       x,y:integer;
       kkw:string;
  begin
       if bnm=5 then kkw:='';
      LL1:=Max;LL2:=Max;
      flag1:=-1;flag2:=-1;
      if A.x<>B.x then K1:=(A.y-B.Y)/(A.x-b.X);
      for i:=0 to Sum-1 do
      begin

        C.x:=npionts[i].x; C.y:=npionts[i].y;

        if ((C.x=A.x)and(C.y=A.y))OR((C.x=B.x)and(C.y=B.y)) or (npionts[i].q=0) then continue;
        if (A.x=B.x)and(B.x=c.x) then continue;
        if (C.x<>B.x) and (A.x<>B.x) and  (K1=(C.y-B.Y)/(C.x-b.X)) then continue;


        H:=ZhongxinPiont(A,B,C);

        LL:=(H.x-Z.x)*(H.x-Z.x)+(H.y-Z.y)*(H.y-Z.y);

        if abs(H.x-Z.x)<0.000001  then
        begin

           if (H.y>Z.y) and (LL1>LL) and (LL>0.000001) then
           begin
             LL1:=LL;flag2:=i;
           end;

           if (H.y<Z.y) and (LL2>LL) and (LL>0.000001) then
           begin
             LL2:=LL;flag1:=i;
           end;
        end ;

      if abs(H.x-Z.x)>0.000001  then
        begin
        if (H.x>Z.x) and (LL1>LL) and (LL>0.000001) then
         begin
           LL1:=LL;flag2:=i;
         end;

        if (H.x<Z.x) and (LL2>LL)and (LL>0.000001)  then
         begin
           LL2:=LL;flag1:=i;
         end;
        end;
      end;

      nm:=0;

      if  flag2<>-1 then
      begin

        C.x:=npionts[flag2].x; C.y:=npionts[flag2].y;
        H:=ZhongxinPiont(A,B,C);

        if direction(A,H,Z)=bz  then
        begin
          npionts[flag2].q:=0; nm:=1;
          if bianjie(H,Z)=1 then result:=flag2
          else
          begin
            npionts[flag2].q:=1;
            result:=KK; kk:=-1;
          end;
        end;

      end;

      if  (flag1<>-1) and (nm=0) then
      begin

        C.x:=npionts[flag1].x; C.y:=npionts[flag1].y;
        H:=ZhongxinPiont(A,B,C);

        if direction(A,H,Z)=bz  then
        begin
          npionts[flag1].q:=0; nm:=1;
          if bianjie(H,Z)=1 then result:=flag1
          else
          begin
            npionts[flag1].q:=1;
            result:=KK; kk:=-1;
          end;
        end;

      end;

      if (nm=0)  then
      begin
        bbianjie(H,Z,A,B);
        result:=KK; kk:=-1;
      end;
  end;


procedure   bbianjie(var H:Tpiont; Z,A,B:Tpiont);
  var   x1,x2,y1,y2,LK,kp,x0,y0:Double;
        pp:integer;
  begin

      Lk:=Max;
      X2:=(A.x+B.X)/2; Y2:=(A.y+B.y)/2;
      pp:=0;

      if  (X2<>Z.X) then
      begin
        X1:=left;Y1:=(Z.y-y2)/(Z.x-x2)*(X1-Z.x)+ Z.y;
        H.x:=X1;H.y:=y1;
        if direction(A,H,Z)=bz then
        begin
          kP:=(X2-x1)*(X2-x1)+(y2-y1)*(y2-y1);
          if kP<lk then
          begin
            lk:=KP;
            x0:=H.x;y0:=H.y;
          end;
        end;
      end;

      if (X2<>Z.X) then
      begin
        X1:=right;Y1:=(Z.y-y2)/(Z.x-x2)*(X1-Z.x)+ Z.y;
        H.x:=X1;H.y:=y1;
        if direction(A,H,Z)=bz then
        begin
          kP:=(X2-x1)*(X2-x1)+(y2-y1)*(y2-y1);
          if kP<lk then
          begin
            lk:=KP;
            x0:=H.x;y0:=H.y;
          end;
        end;
      end;

      if (Z.y<>y2) then
      begin
        y1:=top;x1:=(Z.x-x2)/(Z.y-y2)*(Y1-y2)+ x2;
        H.x:=X1;H.y:=y1;
        if direction(A,H,Z)=bz then
        begin
          kP:=(X2-x1)*(X2-x1)+(y2-y1)*(y2-y1);
          if kP<lk then
          begin
            lk:=KP;
            x0:=H.x;y0:=H.y;
          end;
        end;
      end;

      if (Z.y<>y2) then
      begin
        Y1:=bottom;x1:=(Z.x-x2)/(Z.y-y2)*(Y1-y2)+ x2;
        H.x:=X1;H.y:=y1;
        if direction(A,H,Z)=bz then
        begin
          kP:=(X2-x1)*(X2-x1)+(y2-y1)*(y2-y1);
          if kP<lk then
          begin
            lk:=KP;
            x0:=H.x;y0:=H.y;
          end;
        end;
      end;
      H.x:=X0;H.y:=y0;
  end;


function   bianjie( Var H:Tpiont;Z:Tpiont):integer;
  var   x1,y1:Double;
        pp:integer;
  begin

      if (H.x<left) or (H.x>right) or (H.Y>top) or (h.Y<bottom) then
      begin

        pp:=0;

        if (pp=0) and (H.X<>Z.X) then
        begin
          X1:=left;Y1:=(Z.y-H.y)/(Z.x-H.x)*(X1-H.x)+ H.y;
          if (X1-H.x)*(X1-Z.x)<0 then
          begin
            H.x:=X1;H.y:=y1;pp:=1;
          end;
        end;

        if (pp=0) and (H.X<>Z.X) then
        begin
          X1:=right;Y1:=(Z.y-H.y)/(Z.x-H.x)*(X1-H.x)+ H.y;
          if (X1-H.x)*(X1-Z.x)<0 then
          begin
            H.x:=X1;H.y:=y1;pp:=1;
          end;
        end;

        if (pp=0) then
        begin
          y1:=top;x1:=(Z.x-H.x)/(Z.y-H.y)*(Y1-H.y)+ H.x;
          if (Y1-H.y)*(y1-Z.y)<0 then
          begin
            H.x:=X1;H.y:=y1;pp:=1;
          end;
        end;

        if (pp=0) then
        begin
          y1:=bottom;x1:=(Z.x-H.x)/(Z.y-H.y)*(Y1-H.y)+ H.x;
          if (y1-H.y)*(y1-Z.y)<0 then
          begin
            H.x:=X1;H.y:=y1;pp:=1;
          end;
        end;

        result:=0;

      end
      else  result:=1;

  end;


function direction(A,H,Z:Tpiont):integer;
  var   k,x1,x2,y1,y2,y0:Double;
  begin

      x1:=H.x-A.x;Y1:=H.y-A.y;
      x2:=Z.x-A.x;Y2:=z.y-A.y;

      if X2=0 then
      begin
        if Y2>0
        then
          if x1>0 then result:=1 else  result:=0
        else
          if x1>0 then result:=0 else  result:=1;
      end
      else
      begin

        K:=Y2/X2;y0:=y1-k*X1;

        if  y0>0 then
        begin
          if ((Y2=0)or(Y2>0)) and (x2>0) then result:=0;
          if ((Y2=0)or(Y2>0)) and (x2<0) then result:=1;
          if (Y2<0) and           (x2<0) then result:=1;
          if (Y2<0) and           (x2>0) then result:=0;
        end
        else
        begin
          if ((Y2=0)or(Y2>0)) and (x2>0) then result:=1;
          if ((Y2=0)or(Y2>0)) and (x2<0) then result:=0;
          if (Y2<0)           and (x2<0) then result:=0;
          if (Y2<0)           and (x2>0) then result:=1;
        end;
      end;
  end;

end.
