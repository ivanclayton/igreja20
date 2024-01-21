unit uWrkUnit;

interface

uses Classes,SerialConsts,Registry,Windows,SysUtils,SerialPooler,wtsClient,
     wtsStream,general,prodinfo,wtsmethodview,clientconfigs,Resunit,
     ExecEventoInt, SMCore, SMWtsStorage, uFormDebug, IdBaseComponent,
     IdComponent, IdTCPServer;

type
  PFuncData = ^RFuncData;
  RFuncData = Record
    cod,nom:String;
    id:Integer;
    ok:Boolean;
  End;

  TezResourceLock = class
    {-Encapsulation of a Critical Section}
    protected {private}
      rlCritSect : TRTLCriticalSection;
    protected
    public
      constructor Create;
      destructor Destroy; override;

      procedure Lock;
      procedure Unlock;
  end;
  
  TWrkThread = class;
  TColetor = class
  private
    { Private declarations }
    fId,fLc,fLd,fSq:String;
    fIP,fPeerIP:String;
    fPort: Integer;
    fwk:TWrkThread;
    ped:TwtsMethodView;
    rs:TwtsRecordset;
    fCpg,fVnd:RFuncData;
    fTab,fFil,
    fCli:Integer;
    itm:TItemsCodBarras;
    fTroca,fDesconto:Double;
    fPedOk:Boolean;
    fLock :TezResourceLock;
    fConnID: TIdTCPServerConnection;
    procedure ClearList;
    procedure AddItem(itm:TProdutoCodBarras);
    procedure DeleteItem(itm:TProdutoCodBarras);
    Function  FindItem(itm:TProdutoCodBarras;out qtd:Integer):Boolean;
    Function  ValorPagar:Double;

    procedure FinalizaVenda;
    function GetPedido: String;
  public
    { Public declarations }
    procedure Lock;
    procedure Unlock;

    destructor  Destroy;override;
    constructor Create(const id:String;wrk:TWrkThread);
    Procedure ProcessCMD(Const command,data,seq:String);
    Procedure ReprocCMD;
    property  cod_pedido:String read GetPedido;
    property  LastItem:TItemsCodBarras read itm;
    property  Connection: TIdTCPServerConnection read fConnID write fConnID;
  End;

  TWrkThread = Class(TThread)
  private
    { Private declarations }
    LogFile:TFileStream;
    fSParms: TSerialParams;
    fHandle:TPortID;
    fRestart:Boolean;
    fLastErrorMsg:String;
    fCpgList,
    fFuncList:TStringList;
    fProcList:TStringList;
    fProdInfo:TProdInfo;
    TCPServer:TIdTCPServer;
    rs:TezResourceLock;
    //FWSID : string;
    //FWSTYPEIDX:Integer;
    procedure SetSParms(const Value: TSerialParams);
    procedure SerialOpen;
    procedure SerialClose;
    function  GetCharRdy: Boolean;
    function  GetFunc(Const id: String): RFuncData;
    function  GetCPG(const id:String):RFuncData;
    function  GetCustomer(Const id:String;out Cliente:Integer):String;
    function  GetTabela(const id:String;out Tabela:Integer):Boolean;
    function  GetPI: TProdInfo;
    Function  Devolucao(const Numero:String;out Valor:Double):Boolean;
    procedure ServerExecute(AThread: TIdPeerThread);
    procedure WriteLog(Const msg:String);
    //procedure BuildWSID;
  protected
    { Public declarations }
    procedure Execute; override;
    property  InQue:Boolean Read GetCharRdy;
    property  Funcionario[Const id:String]:RFuncData Read GetFunc;
    property  pi:TProdInfo Read GetPI;
  public
    { Public declarations }
    procedure Write(conn:TIdTCPServerConnection;Const id,cmd,seq,msg:String;dataOk:Boolean);
    property  SerialParams:TSerialParams read fSParms Write SetSParms;
    property  LastError:String read fLastErrorMsg;
    procedure Restart;
    destructor Destroy;override;
  End;

  function  LoadParams(Const Name:String;out Atacado,Varejo:String;out Tcp:Boolean;out IpPort:Integer;out Loja:Integer):TSerialParams;
  procedure RestartWrk;
  procedure ShowDebug;

implementation

var
  wrk:TWrkThread;
  tbA, tbV:String;
  UseTCP:Boolean;
  Loja,IPPort:Integer;
  cs:_RTL_CRITICAL_SECTION;

{ Serial Routines From serlib.dll }

    procedure as_Close(Handle:Integer);stdcall;external 'serlib.dll';
    function as_Write(Handle:Integer;const Buf;Count:Integer):Integer;stdcall;external 'serlib.dll';
    function as_Read(Handle:Integer;var Buf;Count:Integer):Integer;stdcall;external 'serlib.dll';
    function as_ReadEx(Handle:Integer;var Buf;Count:Integer;Terminator:Char):Integer;stdcall;external 'serlib.dll';
    function as_Open(Com,Br,Db,Sb,Pr,Fc:Integer):Integer;stdcall;external 'serlib.dll';
    function as_CharReady(Handle:Integer):Boolean;stdcall;external 'serlib.dll';
    function as_GetChar(Handle:Integer):Char;stdcall;external 'serlib.dll';

{ Internal Functions }

procedure ShowDebug;
Begin
End;

procedure ProtoWrite( conn: TIdTCPServerConnection; const id,cmd,msg,sq: String;dataOk:Boolean=True);
var s: String;
    p: String;
begin
  if conn= nil then
     p:= '@DT'
  else
     p:= '@';

  if DataOk then
     //s:= p + id + cmd + 'OK' + RemoveAcentos(msg)+#13
     s:= p + id + cmd + sq + 'OK' + msg+#13
  else
     //s:= p + id + cmd + 'NOK'+ RemoveAcentos(msg)+#13;
     s:= p + id + cmd + sq + 'NOK'+ msg+#13;
  conn.WriteLn(s);
end;

procedure RestartWrk;
Begin
     wrk.Restart;
End;

function LoadParams(Const Name:String;out Atacado,Varejo:String;out Tcp:Boolean;out IpPort:Integer;out Loja:Integer):TSerialParams;
var r:TRegIniFile;
begin
    with Result do
    begin
         R:=TRegIniFile.Create('Software\Windoor\Cpx2Wts');
         Com := R.ReadInteger(Name,'ComPort',1);
         Br  := R.ReadInteger(Name,'BoudRate',115200);
         Db  := R.ReadInteger(Name,'DataBits',8);
         Sb  := R.ReadInteger(Name,'StopBits',1);
         Pr  := R.ReadInteger(Name,'Parity',pNone);
         Fc  := R.ReadInteger(Name,'FlowControl',fcHardware);
         Atacado := R.ReadString(Name,'Atacado','A');
         Varejo  := R.ReadString(Name,'Varejo','V');
         Tcp := R.ReadBool('Cpx2wts','UseTCP',False);
         IpPort := R.ReadInteger('Cpx2wts','IPPort',1310);
         Loja := R.ReadInteger('Cpx2wts','LojaI',-1);
         R.Free;
    end;
End;

{ TWrkThread }

(*procedure TWrkThread.BuildWSID;
var
   r:TRegIniFile;
   userdata:TwtsTagBag;
   FSessionId:Integer;
begin
     r := TRegIniFile.Create;
     Try
        r.OpenKey('software\millenium',true);
        if r.ReadString('workstation','identification','')='' then
        begin
        end;

        FWSTYPEIDX := r.ReadInteger('workstation','type',0);
        FWSID := r.ReadString('workstation','identification','');
     Finally
        r.free;
     End;
     Randomize;
     FSessionId := Random(Maxint);
     userdata := TwtsTagBag.Create;
     Try
        userdata.Tags['client_username']  := 'CPX2WTS';
        userdata.Tags['client_usergroup'] := 'COLETORES';

        userdata.Tags['client_wsid'] := FWSID;
        userdata.Tags['client_wstype'] := 'lojas';
        userdata.Tags['client_sessionid'] := IntToStr(fSessionId);
     Finally
        Userdata.Free;
     End;
end;*)

destructor TWrkThread.Destroy;
begin
     SerialClose;
     If Assigned(rs) Then
        rs.Free;
     LogFile.Free;
     inherited;
end;

function TWrkThread.Devolucao(const Numero: String;
  out Valor: Double): Boolean;
var v:variant;
begin
     Result := wtsCall('millenium.movimentacao.lista_por_evento',
               ['DOCUMENTO','CANCELADA','GERADOR'],[Numero,False,'C'],v);
     If Result Then
        Valor := VarAsType(v[8],varDouble);
end;

procedure TWrkThread.Execute;
var s,co,cm,sq,da:String;
    ix:Integer;
    col:TColetor;
begin
     //BuildWSID;
     fRestart  := False;
     fFuncList := nil;
     fProcList := TStringList.Create;
     fProcList.Sorted := True;
     fProdInfo := TProdInfo.Create;
     SerialOpen;
     While not Terminated do
     Begin
          Try
             If UseTCP Then
                If not TCPServer.Active Then
                   SerialOpen;

             If fRestart Then
             Begin
                  SerialClose;
                  fRestart := False;
                  SerialOpen;
             End;
             If InQue Then
             Begin
                  Spooler.SerialRead(fHandle,s);
                  If Copy(s,1,3)='@DT' Then
                  Begin
                       WriteLog(s);
                       co := Copy(s,6,2);
                       cm := Copy(s,8,2);
                       sq := Copy(s,10,1);
                       ix := Pos(#13,s);
                       If ix<>0 Then
                          da := Copy(s,11,ix-11)
                       Else
                           da := Copy(s,11,MaxInt);
                       If not fProcList.Find(co,ix) Then
                          ix := fProcList.AddObject(co,TColetor.Create(co,Self));
                       col := TColetor(fProcList.Objects[ix]);
                       col.ProcessCMD(cm,da,sq);
                  End Else
                  If Copy(s,1,3)='@NG' Then
                  Begin
                       co := Copy(s,6,2);
                       If fProcList.Find(co,ix) Then
                          TColetor(fProcList.Objects[ix]).ReprocCMD;
                  End;
             End Else
                 Sleep(100); { Tell to win API to process other messages }
          Except
             On e:Exception do
                OutputDebugString(PChar(e.Message));
          End;
     End;
     For ix:=0 To Pred(fProcList.Count) do
         TColetor(fProcList.Objects[ix]).Free;
     fProcList.Free;
     fProdInfo.Free;
end;

function TWrkThread.GetCharRdy: Boolean;
begin
     If UseTCP Then Result := False
     Else
     Begin                            
         If fHandle.Handle=0 Then SerialOpen;
         If fHandle.Handle=0 Then Result := False
         Else Result := SPooler.InQue[fHandle];
     End;
end;

function TWrkThread.GetCPG(const id: String): RFuncData;
var r:TwtsRecordset;
    i:Integer;
    f:PFuncData;
begin
     If not Assigned(fCPGList) Then
     Begin
          fCPGList := TStringList.Create;
          wtsCallEx('millenium.condicoes_pgto.lista',[''],[],r);
          While not r.Eof do
          Begin
               New(f);
               f.id  := VarAsType(r.FieldValues[0],varInteger);
               f.cod := VarToStr(r.FieldValues[1]);
               f.nom := VarToStr(r.FieldValues[2]);
               fCPGList.AddObject(f.cod,Pointer(f));
               r.Next;
          End;
          r.Free;
          fCPGList.Sort;
     End;
     If fCPGList.Find(id,i) Then
        Result := PFuncData(fCPGList.Objects[i])^
     Else
        Result.nom := '';
end;

function TWrkThread.GetCustomer(const id: String;out Cliente:Integer): String;
var Nome:String;
    Limite:Double;
    v:Variant;
begin
     If wtsCall('millenium.clientes.busca',['COD_CLIENTE'],[id],v) Then
     Begin
          Cliente := v[0];
          Nome    := v[2];
          If wtsCall('millenium.clientes.verifica_cliente',[''],[Cliente],v) Then
          Begin
               If VarToStr(v[0])='' Then
                  Limite := 9999999.99
               Else If v[0]=0 Then
                  Limite := 9999999.99
               Else If VarToStr(v[1])='' Then
                  Limite := StrToFloat(VarToStr(v[0]))
               Else
                  Limite := StrToFloat(VarToStr(v[0]))-StrToFloat(VarToStr(v[1]));
          End Else Limite := 9999999.99;
          Result := Nome + '$' + FloatToStrF(Limite,ffFixed,12,2);
     End Else
         Result := '';
end;

function TWrkThread.GetFunc(Const id: String): RFuncData;
var r:TwtsRecordset;
    i:Integer;
    f:PFuncData;
begin
     If not Assigned(fFuncList) Then
     Begin
          fFuncList := TStringList.Create;
          wtsCallEx('millenium.funcionarios.lista_simples',[''],[],r);
          While not r.Eof do
          Begin
               New(f);
               f.id  := VarAsType(r.FieldValues[0],varInteger);
               f.cod := VarToStr(r.FieldValues[1]);
               f.nom := VarToStr(r.FieldValues[2]);
               f.ok  := True;
               fFuncList.AddObject(f.cod,Pointer(f));
               r.Next;
          End;
          r.Free;
          fFuncList.Sort;
     End;
     If fFuncList.Find(id,i) Then
        Result:= PFuncData(fFuncList.Objects[i])^
     Else
        Result.ok:= False;
end;

function TWrkThread.GetPI: TProdInfo;
begin
     Result := fProdInfo;
end;

function TWrkThread.GetTabela(const id: String;
  out Tabela: Integer): Boolean;
var v:Variant;

    Function TabID:String;
    Begin
         If id='A' Then Result := tbA
            Else Result := tbV;
    End;

begin
     Result := (id = 'A') or (id = 'V');
     If Result Then
        Result := wtsCall('millenium.tabelas_preco.consulta_conversao',[''],[TabID],v);
     If Result Then Tabela := v[0];
end;

procedure TWrkThread.Restart;
begin
     fRestart := True;
end;

procedure TWrkThread.SerialClose;
begin
     if fHandle.Handle<>0 then
     begin
          Spooler.Close(fHandle);
          fHandle.Handle := 0;
     end;
     if Assigned(TCPServer) then
        FreeAndNil(TCPServer);
end;

procedure TWrkThread.SerialOpen;
begin
     If UseTCP Then
     Begin
         SerialClose;
         TCPServer := TIdTCPServer.Create(nil);
         TCPServer.DefaultPort := IPPort;
         TCPServer.OnExecute := ServerExecute;
         TCPServer.Active := True;
     End Else
     Begin
         SPooler.Close(fHandle);
         Try
            fHandle := Spooler.Open(SerialParams);
            Spooler.Terminator := #13;
         Except
         End;
     End;
end;

procedure TWrkThread.ServerExecute(AThread: TIdPeerThread);
var Cmd,da,co,cm,sq:String;
    ix:Integer;
    col:TColetor;
begin
     If not Assigned(rs) Then
        rs := TezResourceLock.Create;
        
     While AThread.Connection.Connected do
     Begin
          //AThread.Connection.ReadTimeout := 10000;

          Try
             cmd:= AThread.Connection.ReadChar;
          except
             AThread.Connection.DisconnectSocket;
             cmd:= '!';
          end;

          If cmd='@' Then
          Begin
               Cmd := '@'+ AThread.Connection.ReadLn(#$D);


               //SetLength(cmd,AThread.Connection.InputBuffer.Size);
               //AThread.Connection.ReadBuffer(cmd[1],Length(cmd));
               //cmd := '@'+cmd;
               //dbg.Add(Cmd);
               //co := Copy(Cmd,2,2);
               //cm := Copy(Cmd,4,2);
               //da := Copy(Cmd,6,MaxInt);

               co := Copy(cmd,2,2);
               cm := Copy(cmd,4,2);
               sq := Copy(cmd,6,1);
               ix := Pos(#13,cmd);
               If ix<>0 Then
                  da := Copy(cmd,7,ix-7)
               Else
                  da := Copy(cmd,7,MaxInt);

               rs.Lock;
               try
                  If not fProcList.Find(co,ix) Then
                  begin
                       col:= TColetor.Create(co,Self);
                       ix := fProcList.AddObject(co,col);
                  end;
                  col:= TColetor(fProcList.Objects[ix]);
               finally
                  rs.Unlock;
               end;

               col.Lock;
               Try
                  col.Connection:= AThread.Connection;
                  Try
                     col.ProcessCMD(cm,da,sq);
                  except
                     ProtoWrite(AThread.Connection,co,cm,'',sq,False);
                  end;
               finally
                  col.Unlock;
               end;


               (*If cm='NG' Then
               Begin
                    If fProcList.Find(co,ix) Then
                       TColetor(fProcList.Objects[ix]).ReprocCMD(AThread.Connection);
               End Else
               Begin

                    If not fProcList.Find(co,ix) Then
                       ix := fProcList.AddObject(co,TColetor.Create(co,Self));
                    col := TColetor(fProcList.Objects[ix]);
                    Try
                       col.ProcessCMD(cm,da,'',AThread.Connection);
                    Except
                       Write(AThread.Connection,co,cm,'','',False);
                    End;
               End;*)

          End;
     End;
end;

procedure TWrkThread.SetSParms(const Value: TSerialParams);
begin
     If not Suspended Then Suspend;
     fSParms := Value;
     Resume;
end;

procedure TWrkThread.Write(conn:TIdTCPServerConnection;const id,cmd,seq,msg: String;dataOk:Boolean);
var s:String;
    p:String;
    sDataOk:String;
begin
     If conn=nil Then
        p := '@DT' Else p := '@';
     If (msg='') and not DataOk Then
        s := p+id+cmd+seq+'NOK'#13
     Else
        s := p+id+cmd+seq+'OK'+msg+#13;

     WriteLog(s);

     If conn=nil Then
        SPooler.SerialWrite(fHandle,s)
     Else
     Begin
          if DataOk then
             sDataOk:= 'True'
          else
             sDataOk:= 'False';

          ProtoWrite(Conn,id,cmd,msg,seq,DataOk);
     End;
end;

procedure TWrkThread.WriteLog(const msg: String);
begin
     EnterCriticalSection(cs);
     Try
        If not Assigned(LogFile) Then
        Begin
             If not FileExists(ExtractFilePath(ParamStr(0))+'\Logcpx.txt') Then
             Begin
                  LogFile := TFileStream.Create(ExtractFilePath(ParamStr(0))+'\Logcpx.txt',fmCreate);
                  logFile.Free;
             End;
             LogFile := TFileStream.Create(ExtractFilePath(ParamStr(0))+'\Logcpx.txt',fmOpenWrite or fmShareDenyNone);
        End;
        LogFile.Write(msg[1],Length(msg));
        If Length(msg)<>0 Then
           LogFile.Write(#13#10,2);
     Finally
        LeaveCriticalSection(cs);
     End;
end;

{ TColetor }

procedure TColetor.AddItem(itm: TProdutoCodBarras);
begin
     If not fPedOk Then fPedOk := True;
     rs.New;
     rs.FieldValues[0] := itm.id;
     rs.FieldValues[1] := StrToInt(itm.Cor);
     rs.FieldValues[2] := StrToInt(itm.Est);
     rs.FieldValues[3] := itm.Tam;
     rs.FieldValues[4] := itm.Qtd;
     rs.FieldValues[5] := itm.value;
     rs.Add;
end;

procedure TColetor.ClearList;
begin
     If not Assigned(rs) Then
     Begin
          rs := TwtsRecordset.CreateFromStream(TMemoryStream.Create);
          rs.Transaction := 'millenium.pedido_venda.produtos';
     End;
     fTroca := 0;
     fDesconto := 0;
     rs.Clear;
     EnterCriticalSection(cs);
     Try
        wrk.pi.Reset;
     Finally
        LeaveCriticalSection(cs);
     End;
end;

constructor TColetor.Create(const id: String;wrk:TWrkThread);
begin
     inherited Create;
     fLock := TezResourceLock.Create;
     fid   := id;
     fwk   := wrk;
     fSq   := '.';
     ped   := TwtsMethodView.Create(nil);
     fFil  := Loja; // SysParam('FILIAL').AsInteger;
end;

procedure TColetor.DeleteItem(itm: TProdutoCodBarras);
begin
     rs.First;
     While not rs.Eof do
     Begin
          If (rs.FieldValues[0]=itm.id) and
             (rs.FieldValues[1]=StrToInt(itm.Cor)) and
             (rs.FieldValues[2]=StrToInt(itm.Est)) and
             (rs.FieldValues[3]=itm.Tam) Then
          Begin
               rs.Delete;
          End;
          rs.Next;
     End;
end;

destructor TColetor.Destroy;
begin
  FreeAndNil(fLock);
  inherited;
end;

procedure TColetor.FinalizaVenda;
var vl:Double;
    ni:Integer;
begin
     If fPedOk Then
     Begin
         fVnd.ok := False;
         fPedOk := False;
         //ped.CallMode := cmAsync;
         ped.Transaction := 'millenium.pedido_venda.inclui';
         ped.ParamsData.Clear;
         ped.ParamsData.New;
         SetToDefaults(ped);

         ni := 0;
         rs.First;
         While not rs.Eof do
         Begin
              Inc(ni);
              rs.SetFieldData(13,Copy(IntToStr(1000+ni),2,3));
              rs.Update;
              rs.Next;
         End;


         With ped.ParamsData do
         Begin
              vl := ValorPagar;
              FieldValues[0]  := cod_pedido;
              FieldValues[1]  := fCli;
              FieldValues[8]  := fFil;
              FieldValues[9]  := Date;
              FieldValues[10] := Date;
              FieldValues[11] := fTab;
              FieldValues[12] := fCpg.id;
              FieldValues[14] := (fDesconto / vl * -100);
              FieldValues[15] := fDesconto * (-1);
              FieldValues[16] := vl;
              FieldValues[20] := rs.Data;
              FieldValues[21] := 'AC';
              FieldValues[22] := 0;
              FieldValues[29] := fVnd.id;
              Add;
         End;

         For ni:=0 To 50 do
         Begin
              Try
                 ped.Refresh;
                 Break;
              Except
                 If ni=50 Then Raise;   
              End;
              ped.ParamsData.FieldValues[0] := cod_pedido;
              ped.ParamsData.Update;
         End;


         With TExecutaEvento.Create(nil) do
         Begin
              Try
                 PrintPedido(ped.Fields[0].AsInteger);
              Except
              End;
              Free;
         End;

         ClearList;
         fVnd.ok := True;
     End;
end;
Function TColetor.FindItem(itm: TProdutoCodBarras;out qtd:Integer):Boolean;
begin
     Result := False;
     qtd := 0;
     rs.First;
     While not rs.Eof do
     Begin
          If (rs.FieldValues[0]=itm.id) and
             (rs.FieldValues[1]=StrToInt(itm.Cor)) and
             (rs.FieldValues[2]=StrToInt(itm.Est)) and
             (rs.FieldValues[3]=itm.Tam) Then
          Begin
               qtd := qtd + rs.FieldValues[4];
               Result := True;
          End;
          rs.Next;
     End;
end;

function TColetor.GetPedido: String;
var v:Variant;
begin
     wtscall('millenium.utils.default',[''],['COD_PEDIDOV'],v);
     Result := v[0];
end;

procedure TColetor.Lock;
begin
    fLock.Lock;
end;

procedure TColetor.ProcessCMD(const command, data, seq:String);
var rs:String;
    vl:Double;
    vqtd,nqtd,oqtd:Integer;
    fUserList: TUserList;

    Function FloatToStrP(valor:Double):String;
    var p:Integer;
    Begin
         Result := FloatToStrF(Valor,ffFixed,12,2);
         p := Pos(',',Result);
         If p>0 Then Result[p] := '.';
    End;

    Function SameCMD:Boolean;
    Begin
         Result := ((fLc=command) and (fSq=seq));
    End;
begin
     If Assigned(fConnId) Then
     With fConnId.Socket.Binding do
     begin
          fip    := IP;
          fPort  := Port;
          fPeerIP:= PeerIP;
     end;


     fLd := data;
     If command='AT' then
     Begin
          wrk.Write(Connection,fId,'AT',seq,'',True);
     End Else
     If command='CO' Then  // Código do Operador
     Begin
        fVnd := wrk.GetFunc(data);
        wrk.Write(Connection,fId,'NO',seq,fVnd.nom,fVnd.ok);
     End Else If command='SO' Then  // Senha do Operador
     Begin
           fUserList := TUserList.Create;
           Try
              fUserList.Storage := TwtsStorage.Create;

              try
                 fUserList.DoLogin(fVnd.nom,data,ltReadWrite);
                 wrk.Write(Connection,fId,'LO',seq,'',True);
              except
                 wrk.Write(Connection,fId,'LO',seq,'',False);
              end;
           Finally
              fUserList.Free;
           End;
     End Else
     If command='CC' Then // Cliente
     Begin
          rs := wrk.GetCustomer(data,fCli);
          wrk.Write(Connection,fId,'NC',seq,rs,(fVnd.ok and (rs<>'')));
     End Else
     If command='TV' Then // Tipo de Venda
     Begin
          wrk.Write(Connection,fId,'AV',seq,'',(fVnd.ok and wrk.GetTabela(data,fTab)));
          ClearList;
     End else
     If command='CP' Then // Codigo do Produto
     Begin
          Try
             If not fVnd.ok Then Raise Exception.Create('');
             SetLength(itm,0);
             LeBarraProduto(data,'AC',itm,wrk.pi);
             If High(itm)<0 Then
                wrk.Write(Connection,fId,'DP',seq,'',False)
             Else
             Begin
                  Lock;
                  Try
                     itm[0].value := wrk.pi.Preco(fTab,StrToInt(itm[0].Cor),
                                     StrToInt(itm[0].Est),itm[0].Tam,
                                     False, False, True);
                  Finally
                     Unlock;
                  End;
                  wrk.Write(Connection,fId,'DP',seq,Copy(wrk.pi.Descricao,1,40)+'$'+
                  FloatToStrP(itm[0].value)+'%0',True);
             End;
          Except
             wrk.Write(Connection,fId,'DP',seq,'',False);
          End;
     End Else
     If command='IP' Then // Insere Produto
     Begin
          itm[0].Qtd := StrToIntDef(Copy(data,Pos('%',data)+1,MaxInt),0);
          If not SameCMD and (itm[0].Qtd<>0) Then
             AddItem(itm[0]);
          wrk.Write(Connection,fId,'AP',seq,'',fVnd.ok);
     End Else
     If command='AQ' Then // Altera Quantidade do Produto
     Begin
          Try
             If not fVnd.ok Then Raise Exception.Create('');
             SetLength(itm,0);
             LeBarraProduto(Copy(data,1,Pred(Pos('%',data))),'AC',itm,wrk.pi);
             rs   := Copy(data,Succ(Pos('%',data)),MaxInt);
             oqtd := StrToIntDef(Copy(rs,1,Pred(Pos('%',rs))),0);
             nqtd := StrToIntDef(Copy(rs,Succ(Pos('%',rs)),MaxInt),0);
             If (High(itm)<0) or not FindItem(itm[0],vqtd)
                or ((oqtd<>0) and (vqtd<>oqtd)) Then
                wrk.Write(Connection,fId,'AP',seq,'',False)
             Else
             Begin
                  If oqtd<>0 Then
                     DeleteItem(itm[0]);
                  itm[0].Qtd := nqtd;
                  If itm[0].Qtd<>0 Then
                  Begin
                       itm[0].value := wrk.pi.Preco(fTab,StrToInt(itm[0].Cor),
                                       StrToInt(itm[0].Est),itm[0].Tam);
                       AddItem(itm[0]);
                  End;
                  wrk.Write(Connection,fId,'AP',seq,'',fVnd.ok);
             End;
          Except
            wrk.Write(Connection,fId,'DP',seq,'',False);
          End;
     End Else
     If command='DT' Then // Definição de Troca
     Begin
          If wrk.Devolucao(data,fTroca) Then
             wrk.Write(Connection,fId,'VT',seq,FloatToStrP(ValorPagar),fVnd.ok)
          Else
             wrk.Write(Connection,fId,'VT',seq,'',False);
     End Else
     If command='DD' Then // Definição de Desconto
     Begin
          vl := ValorPagar;
          If Copy(data,1,1)='$' Then
             fDesconto := StrToFloat(ClearFloat(Copy(data,2,MaxInt)))/100
          Else
             fDesconto := (vl * StrToFloat(ClearFloat(Copy(data,2,MaxInt)))/100)/100;
          wrk.Write(Connection,fId,'VD',seq,FloatToStrP((vl - fDesconto)),fVnd.ok);
     End Else
     If command='CG' Then // Forma de Pagto
     Begin
          fCpg := wrk.GetCPG(data);
          wrk.Write(Connection,fId,'DG',seq,fCpg.nom,(fVnd.ok and (fCpg.nom<>'')));
     End Else
     If command='FV' Then
     Begin
          If not SameCMD THen
             FinalizaVenda;
          wrk.Write(Connection,fId,'TV',seq,'',fVnd.ok);
     End Else
     If command='CV' Then
     Begin
          ClearList;
          wrk.Write(Connection,fId,'EV',seq,'',True);
     End else
     If command='IS' Then // VENDEDOR DISPONIVEL
     Begin
       wrk.Write(Connection,fId,'FS',seq,'',True);
     End else
     If command='II' Then // VENDEDOR INDISPONIVEL
     Begin
       wrk.Write(Connection,fId,'FI',seq,'',True);
     End;
     fLc := command;
     fSq := seq;
end;

procedure TColetor.ReprocCMD;
begin
     ProcessCMD(fLc,fLd,fSq);
end;

procedure TColetor.Unlock;
begin
    fLock.UnLock;
end;

function TColetor.ValorPagar: Double;
begin
     Result := 0;
     rs.First;
     While not rs.Eof do
     Begin
          Result := Result + (rs.FieldValues[4] * rs.FieldValues[5]);
          rs.Next;
     End;
     Result := Result - fTroca;
end;

{===TezResourceLock==================================================}
constructor TezResourceLock.Create;
begin
  inherited Create;
  InitializeCriticalSection(rlCritSect);
end;
{--------}
destructor TezResourceLock.Destroy;
begin
  DeleteCriticalSection(rlCritSect);
  inherited Destroy;
end;
{--------}
procedure TezResourceLock.Lock;
begin
  EnterCriticalSection(rlCritSect);
end;
{--------}
procedure TezResourceLock.Unlock;
begin
  LeaveCriticalSection(rlCritSect);
end;
{====================================================================}

initialization
Begin
     wrk := TWrkThread.Create(True);
     wrk.FreeOnTerminate := True;
     wrk.SerialParams := LoadParams('Cpx2wts',tbA, tbV, UseTCP, IPPort,Loja);
     InitializeCriticalSection(cs);
End;

Finalization
    wrk.Terminate;
    DeleteCriticalSection(cs);

end.
