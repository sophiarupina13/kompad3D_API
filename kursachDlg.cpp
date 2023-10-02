#include "stdafx.h"
#include "pch.h"
#include "WSCAD.h"
#include "WSCADDlg.h"
#include "kursach.h"
#include "afxdialogex.h"
#include "math.h"
#include  <cmath>
#include "C:\Users\marin\kompas\SDK\Include\ksConstants.h"
#include "C:\Users\marin\kompas\SDK\Include\ksConstants3D.h"
#include <atlsafe.h>
#include <comutil.h>
#include "kursachDlg.h"

#define PI 4*atan(1)
#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#pragma once
#import "C:\Users\marin\kompas\SDK\lib\kAPI5.tlb"

using namespace Kompas6API5;
KompasObjectPtr pKompasApp5;
ksPartPtr pPart;
ksDocument3DPtr pDoc;



// Диалоговое окно CAboutDlg используется для описания сведений о приложении

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// Данные диалогового окна
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // поддержка DDX/DDV

	// Реализация
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// Диалоговое окно CkursachDlg



CkursachDlg::CkursachDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_KURSACH_DIALOG, pParent)
	, D_tot(93)
	, D_hole(88)
	, H_tot(93)
	, H2(55)
	, H1(38)
	, H_step(3)
	, d_fmhs(118)
	, d_ofh(5)
	, sR(100)
	, R_arc(5)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CkursachDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT3, H_tot);
	DDV_MinMaxDouble(pDX, H_tot, 1, 1000);
	DDX_Text(pDX, IDC_EDIT4, H2);
	DDV_MinMaxDouble(pDX, H2, H1 + 1, H_tot - 1);
	DDX_Text(pDX, IDC_EDIT5, H1);
	DDV_MinMaxDouble(pDX, H1, H_step + 1, H2 - 1);
	DDX_Text(pDX, IDC_EDIT6, H_step);
	DDV_MinMaxDouble(pDX, H_step, 0, H1 - 1);
	DDX_Text(pDX, IDC_EDIT10, sR);
	DDV_MinMaxDouble(pDX, sR, 1, 1000);
	DDX_Text(pDX, IDC_EDIT1, D_tot);
	DDV_MinMaxDouble(pDX, D_tot, sR - 10, sR - 7);
	DDX_Text(pDX, IDC_EDIT2, D_hole);
	DDV_MinMaxDouble(pDX, D_hole, D_tot - 5, D_tot - 1);
	DDX_Text(pDX, IDC_EDIT7, d_fmhs);
	DDV_MinMaxDouble(pDX, d_fmhs, sR + 8, sR + 30);
	DDX_Text(pDX, IDC_EDIT9, d_ofh);
	DDV_MinMaxDouble(pDX, d_ofh, 1, 7);
	DDX_Text(pDX, IDC_EDIT11, R_arc);
	DDV_MinMaxDouble(pDX, R_arc, 0, sR / 2);
}

BEGIN_MESSAGE_MAP(CkursachDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CkursachDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// Обработчики сообщений CkursachDlg

BOOL CkursachDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Добавление пункта "О программе..." в системное меню.

	// IDM_ABOUTBOX должен быть в пределах системной команды.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Задает значок для этого диалогового окна.  Среда делает это автоматически,
	//  если главное окно приложения не является диалоговым
	SetIcon(m_hIcon, TRUE);			// Крупный значок
	SetIcon(m_hIcon, FALSE);		// Мелкий значок

	// TODO: добавьте дополнительную инициализацию

	return TRUE;  // возврат значения TRUE, если фокус не передан элементу управления
}

void CkursachDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// При добавлении кнопки свертывания в диалоговое окно нужно воспользоваться приведенным ниже кодом,
//  чтобы нарисовать значок.  Для приложений MFC, использующих модель документов или представлений,
//  это автоматически выполняется рабочей областью.

void CkursachDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // контекст устройства для рисования

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Выравнивание значка по центру клиентского прямоугольника
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Нарисуйте значок
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}


bool CkursachDlg::CheckData()
{
	if (!UpdateData())
		return false;
	if (sR - d_fmhs < 8 && sR - d_fmhs > 20)
		return false;
	return true;
}

// Система вызывает эту функцию для получения отображения курсора при перемещении
//  свернутого окна.
HCURSOR CkursachDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CkursachDlg::OnBnClickedButton1()
{
	// TODO: добавьте свой код обработчика уведомлений
	BeginWaitCursor();

	if (!CheckData())
		return;

	CComPtr<IUnknown> pKompasAppUnk = nullptr;
	if (!pKompasApp5)
	{
		// Получаем CLSID для Компас
		CLSID InvAppClsid;
		HRESULT hRes = CLSIDFromProgID(L"Kompas.Application.5", &InvAppClsid);
		if (FAILED(hRes)) {
			pKompasApp5 = nullptr;
			return;
		}

		// Проверяем есть ли запущенный экземпляр Компас
		//если есть получаем IUnknown
		hRes = ::GetActiveObject(InvAppClsid, NULL, &pKompasAppUnk);
		if (FAILED(hRes)) {
			// Приходится запускать Компас самим так как работающего нет
			// Также получаем IUnknown для только что запущенного приложения Компас
			TRACE(L"Could not get hold of an active Inventor, will start a new session\n");
			hRes = CoCreateInstance(InvAppClsid, NULL, CLSCTX_LOCAL_SERVER, __uuidof(IUnknown), (void**)&pKompasAppUnk);
			if (FAILED(hRes)) {
				pKompasApp5 = nullptr;
				return;
			}
		}

		// Получаем интерфейс приложения Компас
		hRes = pKompasAppUnk->QueryInterface(__uuidof(KompasObject), (void**)&pKompasApp5);
		if (FAILED(hRes)) {
			return;
		}
	}

	// делаем Компас видимым
	pKompasApp5->Visible = true;

	pDoc = pKompasApp5->Document3D();

	pDoc->Create(false, true);

	pPart = pDoc->GetPart(pTop_Part);

	ksEntityPtr pSketch = pPart->NewEntity(o3d_sketch);

	ksSketchDefinitionPtr pSketchDef = pSketch->GetDefinition();

	pSketchDef->SetPlane(pPart->GetDefaultEntity(o3d_planeXOY));

	pSketch->Create();

	ksDocument2DPtr p2DDoc = pSketchDef->BeginEdit();

	p2DDoc->ksLineSeg((sR / 2) * (-1), sR / 2 - 2 * R_arc, sR / -2, (sR / 2 - 2 * R_arc) * (-1), 1);
	p2DDoc->ksLineSeg(sR / 2 - 2 * R_arc, sR / 2, (sR / 2 - 2 * R_arc) * (-1), sR / 2, 1);
	p2DDoc->ksLineSeg(sR / 2 - 2 * R_arc, (sR / 2) * (-1), (sR / 2 - 2 * R_arc) * (-1), (sR / 2) * (-1), 1);
	p2DDoc->ksLineSeg(sR / 2, sR / 2 - 2 * R_arc, sR / 2, (sR / 2 - 2 * R_arc) * (-1), 1);
	p2DDoc->ksArcByPoint((sR / 2 - 2 * R_arc) * (-1), sR / 2 - 2 * R_arc, R_arc * 2, (sR / 2) * (-1), sR / 2 - 2 * R_arc, (sR / 2 - 2 * R_arc) * (-1), sR / 2, -1, 1);
	p2DDoc->ksArcByPoint(sR / 2 - 2 * R_arc, sR / 2 - 2 * R_arc, R_arc * 2, sR / 2 - 2 * R_arc, sR / 2, sR / 2 - 1, sR / 2 - 2 * R_arc, -1, 1);
	p2DDoc->ksArcByPoint(sR / 2 - 2 * R_arc, (sR / 2 - 2 * R_arc) * (-1), R_arc * 2, sR / 2, (sR / 2 - 2 * R_arc) * (-1), sR / 2 - 2 * R_arc, (sR / 2) * (-1), -1, 1);
	p2DDoc->ksArcByPoint((sR / 2 - 2 * R_arc) * (-1), (sR / 2 - 2 * R_arc) * (-1), R_arc * 2, (sR / 2 - 2 * R_arc) * (-1), (sR / 2) * (-1), (sR / 2) * (-1), (sR / 2 - 2 * R_arc) * (-1), -1, 1);

	pSketchDef->EndEdit();

	ksEntityPtr pExtrude = pPart->NewEntity(o3d_baseExtrusion);
	ksBaseExtrusionDefinitionPtr pExDef = pExtrude->GetDefinition();
	pExDef->directionType = dtNormal;
	pExDef->SetSideParam(true, etBlind, 3, 0, true);
	pExDef->SetSketch(pSketch);
	pExtrude->Create();

	ksEntityPtr pSketch1 = pPart->NewEntity(o3d_sketch);
	pSketchDef = pSketch1->GetDefinition();
	pSketchDef->SetPlane(pPart->GetDefaultEntity(o3d_planeYOZ));
	pSketch1->Create();

	p2DDoc = pSketchDef->BeginEdit();
	p2DDoc->ksLineSeg(-H_tot, D_tot / 2,-H_tot, D_hole / 2, 1);
	p2DDoc->ksLineSeg(-H2, D_tot / 2, -H_tot, D_tot / 2, 1);
	p2DDoc->ksLineSeg(-H1, sR / 2 + 3, -H2, D_tot / 2, 1);
	p2DDoc->ksLineSeg(-H_step, sR / 2, -H1, sR / 2 + 3, 1);
	p2DDoc->ksLineSeg(-H_step, D_hole / 2, -H_step, sR / 2, 1);
	p2DDoc->ksLineSeg(-H_step, D_hole / 2, -H_tot, D_hole / 2, 1);
	p2DDoc->ksLineSeg(-H_step, 0, -H_tot, 0, 3);
	pSketchDef->EndEdit();

	ksEntityPtr pRotate = pPart->NewEntity(o3d_bossRotated);

	ksBossRotatedDefinitionPtr pRotDef = pRotate->GetDefinition();
	pRotDef->SetSketch(pSketch1);
	pRotDef->SetSideParam(TRUE, 360);
	pRotate->Create();

	ksEntityPtr pSketch2 = pPart->NewEntity(o3d_sketch);
	pSketchDef = pSketch2->GetDefinition();
	pSketchDef->SetPlane(pPart->GetDefaultEntity(o3d_planeXOY));
	pSketch2->Create();
	p2DDoc = pSketchDef->BeginEdit();
	p2DDoc->ksCircle(-round(sqrt(pow(d_fmhs / 2, 2) / 2)), -round(sqrt(pow(d_fmhs / 2, 2) / 2)), d_ofh / 2, 1);
	p2DDoc->ksCircle(-round(sqrt(pow(d_fmhs / 2, 2) / 2)), round(sqrt(pow(d_fmhs / 2, 2) / 2)), d_ofh / 2, 1);
	p2DDoc->ksCircle(round(sqrt(pow(d_fmhs / 2, 2) / 2)), round(sqrt(pow(d_fmhs / 2, 2) / 2)), d_ofh / 2, 1);
	p2DDoc->ksCircle(round(sqrt(pow(d_fmhs / 2, 2) / 2)), -round(sqrt(pow(d_fmhs / 2, 2) / 2)), d_ofh / 2, 1);
	p2DDoc->ksCircle(0, 0, D_hole / 2, 1);
	pSketchDef->EndEdit();

	ksEntityPtr pExtrude1 = pPart->NewEntity(o3d_cutExtrusion);
	ksCutExtrusionDefinitionPtr pExDef1 = pExtrude1->GetDefinition();
	pExDef1->directionType = dtReverse;
	pExDef1->SetSketch(pSketch2);
	pExDef1->SetSideParam(true, etBlind, 100, 0, true);
	pExtrude1->Create();

	ksEntityPtr pPlane = pPart->NewEntity(o3d_planeOffset);
	ksPlaneOffsetDefinitionPtr pPlaneDef = pPlane->GetDefinition();
	pPlaneDef->direction = false;
	pPlaneDef->offset = sR / 2;
	pPlaneDef->SetPlane(pPart->GetDefaultEntity(o3d_planeYOZ));
	pPlane->Create();

	ksEntityPtr pSketch3 = pPart->NewEntity(o3d_sketch);
	pSketchDef = pSketch3->GetDefinition();
	pSketchDef->SetPlane(pPlane);
	pSketch3->Create();
	p2DDoc = pSketchDef->BeginEdit();
	p2DDoc->ksLineSeg(0, sR / 2, -sR, sR / 2, 1);
	p2DDoc->ksLineSeg(-sR, sR / 2, -sR, -sR / 2, 1);
	p2DDoc->ksLineSeg(-sR, -sR / 2, 0, -sR / 2, 1);
	p2DDoc->ksLineSeg(0, -sR / 2, 0, sR / 2, 1);
	pSketchDef->EndEdit();

	ksEntityPtr pExtrude2 = pPart->NewEntity(o3d_cutExtrusion);
	pExDef1 = pExtrude2->GetDefinition();
	pExDef1->directionType = dtNormal;
	pExDef1->SetSideParam(true, etBlind, 10, 0, false);
	pExDef1->SetSketch(pSketch3);
	pExtrude2->Create();

	ksEntityCollectionPtr fledges = pPart->EntityCollection(o3d_edge);
	ksEntityPtr pFillet = pPart->NewEntity(o3d_fillet);

	ksFilletDefinitionPtr pFilletDef = pFillet->GetDefinition();
	pFilletDef->radius = 0.2f;
	ksEntityCollectionPtr fl = pFilletDef->array();
	fl->Clear();

	ksEntityPtr pCircCopy = pPart->NewEntity(o3d_circularCopy);
	ksCircularCopyDefinitionPtr pCircDef = pCircCopy->GetDefinition();
	pCircDef->Putcount2(4);
	pCircDef->SetAxis(pPart->GetDefaultEntity(o3d_axisOZ));
	fl = pCircDef->GetOperationArray();
	fl->Clear();
	fl->Add(pExtrude2);
	pCircCopy->Create();

}
