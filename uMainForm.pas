unit uMainForm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Registry, uWrkUnit, Menus, SerialConsts, CoolTrayIcon,
  General, ComCtrls, wtsStream, wtsClient;

type
  TForm1 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    PopupMenu1: TPopupMenu;
    Configurar1: TMenuItem;
    N1: TMenuItem;
    Sair1: TMenuItem;
    CoolTrayIcon1: TCoolTrayIcon;
    Button1: TButton;
    Debug1: TMenuItem;
    Label7: TLabel;
    Edit1: TEdit;
    Label8: TLabel;
    Edit2: TEdit;
    PageControl1: TPageControl;
    RS232: TTabSheet;
    TCPIP: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    ComboBox4: TComboBox;
    ComboBox5: TComboBox;
    ComboBox6: TComboBox;
    Label9: TLabel;
    Edit3: TEdit;
    Label10: TLabel;
    ComboBox7: TComboBox;
    procedure BitBtn1Click(Sender: TObject);
    procedure Sair1Click(Sender: TObject);
    procedure Configurar1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Debug1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

Const Name='Cpx2wts';

procedure TForm1.BitBtn1Click(Sender: TObject);
var r:TRegIniFile;
begin
     R:=TRegIniFile.Create('Software\Windoor\Cpx2Wts');
     R.WriteBool('Cpx2wts','UseTCP',(PageControl1.ActivePage=TCPIP));
     R.WriteInteger('Cpx2wts','IPPort',StrToInt(Edit3.Text));
     R.WriteInteger('Cpx2wts','ComPort',ComboBox1.ItemIndex+1);
     R.WriteInteger('Cpx2wts','BoudRate',StrToInt(ComboBox2.Text));
     R.WriteInteger('Cpx2wts','DataBits',StrToInt(ComboBox3.Text));
     R.WriteString('Cpx2wts','StopBits',ComboBox4.Text);
     R.WriteInteger('Cpx2wts','Parity',ComboBox5.ItemIndex);
     R.WriteInteger('Cpx2wts','FlowControl',ComboBox6.ItemIndex);
     R.WriteString('Cpx2wts','Atacado',Edit1.Text);
     R.WriteString('Cpx2wts','Varejo',Edit2.Text);
     R.WriteInteger('Cpx2wts','LojaI',Integer(ComboBox7.Items.Objects[ComboBox7.ItemIndex]));
     R.WriteString('Cpx2wts','LojaS',ComboBox7.Text);
     R.Free;
     ShowMessage('Reinicie o Aplicativo');
     Application.Minimize;
end;

procedure TForm1.Sair1Click(Sender: TObject);
begin
     Application.Terminate;
end;

procedure TForm1.Configurar1Click(Sender: TObject);
var tbA, tbV:String;
    UseTCP:Boolean;
    x, Loja, IPPort:Integer;
begin
     With LoadParams('Cpx2wts',tbA, tbV, UseTCP, IpPort,Loja) do
     Begin
          If UseTCP Then
             PageControl1.ActivePage := TCPIP
          Else PageControl1.ActivePage := RS232;
          Edit3.Text := IntToStr(IpPort);
          ComboBox1.ItemIndex := Com-1;
          ComboBox2.ItemIndex := ComboBox2.Items.IndexOf(IntToStr(br));
          ComboBox3.ItemIndex := ComboBox3.Items.IndexOf(IntToStr(Db));
          ComboBox4.ItemIndex := ComboBox4.Items.IndexOf(IntToStr(sb));
          ComboBox5.ItemIndex := Pr;
          ComboBox6.ItemIndex := fc;

          FormActivate(nil);
          For x:=0 To Pred(ComboBox7.Items.Count) do
          Begin
               If Integer(ComboBox7.Items.Objects[x])=Loja Then
               Begin
                    ComboBox7.ItemIndex := x;
                    Break;
               End;
          End;

          Edit1.Text := tbA;
          Edit2.Text := tbV;
     End;
     Application.Restore;
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
     Application.Minimize;
end;

procedure TForm1.Button1Click(Sender: TObject);
var col:TColetor;
begin
     col := TColetor.Create('Z',nil);  //7153-8710
     col.ProcessCMD('AT','','');
     col.ProcessCMD('CO','00137','');
     col.ProcessCMD('CC','1055004','');
     col.ProcessCMD('CP','0002001','');
     ShowMessage('4000030 ' + col.LastItem[0].dPro);
     col.ProcessCMD('CP','7896617400021','');
     ShowMessage('7896617400021 ' + col.LastItem[0].dPro);
     col.ProcessCMD('CG','1','');
     col.ProcessCMD('FV','','');
     col.Free;
end;

procedure TForm1.Debug1Click(Sender: TObject);
begin
     ShowDebug;
end;

procedure TForm1.FormActivate(Sender: TObject);
var r:TwtsRecordset;
begin
     If ComboBox7.Items.Count=0 Then
     Begin
          Try
            wtsCallEx('millenium.filiais.lista_simples',[''],[],r);
            Try
               While not r.Eof do
               Begin
                    ComboBox7.Items.AddObject(VarToStr(r.FieldValues[0]),Pointer(StrToIntDef(VarToStr(r.FieldValues[2]),-1)));
                    r.Next;
               End;
            Finally
               r.Free;
            End;
          Except
          End;
     End;
end;

end.
