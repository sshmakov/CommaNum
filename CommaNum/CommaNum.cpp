// CommaNum.cpp : Defines the entry point for the application.
//


#include "stdafx.h"
#include "CommaNum.h"

#define MAX_LOADSTRING 100

#include <Shellapi.h>
#include <Commctrl.h>

#define CHECK_MIN false
#define MIN_NUMBER 0

// Global Variables:
HWND hWnd;
HINSTANCE hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name


NOTIFYICONDATA nid = {};

#define ICON_NOTIFY (WM_USER+6767)

UINT CF_CSV,CF_HTML;

void AddIcon(HWND win)
{
	memset(&nid,0,sizeof(nid));
	nid.cbSize = sizeof(nid);
	nid.hWnd = win;
	nid.uID = 1;
	nid.uFlags = NIF_ICON | NIF_MESSAGE;
	nid.uCallbackMessage  = ICON_NOTIFY;
	nid.hIcon = 0; // !!
	//nid.szTip = "CommaNum";
	nid.dwState = 0;
	nid.dwStateMask = 0;
	//nid.szInfo = "Info";
	//nid.uTimeout
	//nid.uVersion
	//nid.szInfoTitle = "Title";
	nid.dwInfoFlags = NIIF_NONE;
	//nid.hBalloonIcon = 0;
	nid.hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_SMALL));
	Shell_NotifyIcon(NIM_ADD,&nid);
	CF_CSV=RegisterClipboardFormatA("Csv");
	CF_HTML=RegisterClipboardFormatA("HTML Format");
}

void RemoveIcon(HWND win)
{
	Shell_NotifyIcon(NIM_DELETE,&nid);
	//CF_CSV=RegisterClipboardFormatA("Csv");
}




// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

 	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_COMMANUM, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}
	AddIcon(hWnd);
	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_COMMANUM));

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_COMMANUM));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_COMMANUM);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{

   hInst = hInstance; // Store instance handle in our global variable

   hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

   if (!hWnd)
   {
      return FALSE;
   }

   //ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

void ClipToList(UINT frm)
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(frm); 
	if (hglb != NULL) 
	{ 
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			LPSTR p = lpstr;
			bool first=true;
			bool digit=false;
			int len = strlen((const char *)p);
			LPSTR buf = (LPSTR)malloc(len+1024);
			LPSTR dest = buf;
			while(*p)
			{
				if(*p >= '0' && *p <='9')
				{
					if(!digit && !first)
					{
						*dest++ = ',';
						*dest++ = ' ';
					}
					*dest++ = *p++;
					digit = true;
				}
				else
				{
					if(digit)
					{
						first = false;
						digit = false;
					}
					p++;
				}
			}
			*dest=0;
			
			len=strlen(buf);
			HGLOBAL   hnew = GlobalAlloc(GMEM_MOVEABLE, len+1);
			LPSTR glbuf = (LPSTR)GlobalLock(hnew);
			strcpy(glbuf,buf);
			free(buf);
			GlobalUnlock(hnew);
			GlobalUnlock(hglb); 
			EmptyClipboard();
			SetClipboardData(CF_TEXT,hnew);

			// Call the application-defined ReplaceSelection 
			// function to insert the text and repaint the 
			// window. 

			//ReplaceSelection(hwndSelected, pbox, lptstr); 
		} 
	} 
	CloseClipboard(); 
}

int compare(const void *p1, const void *p2)
{
	int i = *(const int*)p1;
	int j = *(const int*)p2;
	return i == j ? 0 : i<j ? -1 : 1;
}

void freemem(void *p)
{
	if(p)
		free(p);
}

void ClipToUnique(UINT frm)
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(frm); 
	if (hglb != NULL) 
	{ 
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			LPSTR p = lpstr;
			bool digit=false;
			int len = strlen((const char *)p);
			int *arr=(int*)malloc(len*sizeof(int));
			memset(arr,0,len*sizeof(int));
			int cur=-1;
			while(*p)
			{
				if(*p >= '0' && *p <='9')
				{
					if(!digit)
						cur++;
					arr[cur]=arr[cur]*10+(*p-'0');
					digit = true;
				}
				else
				{
					digit = false;
				}
				p++;
			}
			qsort(arr,cur+1,sizeof(int),compare);
			LPSTR buf =(LPSTR)malloc(len);
			int prev=0;
			LPSTR dest=buf;
			bool first=true;
			for(int i=0; i<=cur; i++)
			{
				if(CHECK_MIN && arr[i] < MIN_NUMBER)
					continue;
				if(arr[i] == prev)
					continue;
				prev=arr[i];
				if(!first)
				{
					*dest++ = ',';
					*dest++ = ' ';
				}
				first = false;
				_itoa(arr[i],dest,10);
				dest+=strlen(dest);
			}
			*dest=0;
			freemem(arr);
			len=strlen(buf);
			HGLOBAL   hnew = GlobalAlloc(GMEM_MOVEABLE, len+1);
			LPSTR glbuf = (LPSTR)GlobalLock(hnew);
			strcpy(glbuf,buf);
			freemem(buf);
			GlobalUnlock(hnew);
			GlobalUnlock(hglb); 
			EmptyClipboard();
			SetClipboardData(CF_TEXT,hnew);

			// Call the application-defined ReplaceSelection 
			// function to insert the text and repaint the 
			// window. 

			//ReplaceSelection(hwndSelected, pbox, lptstr); 
		} 
	} 
	CloseClipboard(); 
}

// return elements count
int ParseToIntArr(LPSTR text, int *arr)
{
	LPSTR p=text;
	bool digit=false;
	int cur=-1;
	while(*p)
	{
		if(*p >= '0' && *p <='9')
		{
			if(!digit)
				arr[++cur]=0;
			arr[cur]=arr[cur]*10+(*p-'0');
			digit = true;
		}
		else
		{
			digit = false;
		}
		p++;
	}
	return cur+1;
}


LPSTR ArrToUnique(int *arr, int count, bool dies=false)
{
	if(!count)
		return NULL;
	qsort(arr,count,sizeof(int),compare);
	LPSTR buf =(LPSTR)malloc(20*count*sizeof(buf[0]));
	int prev=0;
	LPSTR dest=buf;
	bool first=true;
	for(int i=0; i<count; i++)
	{
		if(CHECK_MIN && arr[i] < MIN_NUMBER)
			continue;
		if(arr[i] == prev)
			continue;
		prev=arr[i];
		if(!first)
		{
			*dest++ = ',';
			*dest++ = ' ';
		}
		first = false;
		if(dies)
			*dest++ = '#';
		_itoa(arr[i],dest,10);
		dest+=strlen(dest);
	}
	*dest=0;
	return buf;
}

LPSTR GetUnique(LPSTR text)
{
	LPSTR p = text;
	int len = strlen((const char *)p);
	int *arr=(int*)malloc(len*sizeof(int));
	memset(arr,0,len*sizeof(int));
	int count = ParseToIntArr(text, arr);
	if(!count)
		return NULL;
	/*
	int cur=-1;
	while(*p)
	{
		if(*p >= '0' && *p <='9')
		{
			if(!digit)
				cur++;
			arr[cur]=arr[cur]*10+(*p-'0');
			digit = true;
		}
		else
		{
			digit = false;
		}
		p++;
	}
	*/
	qsort(arr,count,sizeof(int),compare);
	LPSTR buf =(LPSTR)malloc(15*count*sizeof(buf[0]));
	int prev=0;
	LPSTR dest=buf;
	bool first=true;
	for(int i=0; i<count; i++)
	{
		if(CHECK_MIN && arr[i] < MIN_NUMBER)
			continue;
		if(arr[i] == prev)
			continue;
		prev=arr[i];
		if(!first)
		{
			*dest++ = ',';
			*dest++ = ' ';
		}
		first = false;
		_itoa(arr[i],dest,10);
		dest+=strlen(dest);
	}
	*dest=0;
	freemem(arr);
	return buf;
}

LPSTR memory=NULL;

void SetToClip(LPSTR text, UINT format=CF_TEXT, bool clear=true)
{
	if(text==NULL)
	{
		EmptyClipboard();
		return;
	}
	HGLOBAL hnew = GlobalAlloc(GMEM_MOVEABLE, strlen(text)+1);
	LPSTR glbuf = (LPSTR)GlobalLock(hnew);
	strcpy(glbuf,text);
	GlobalUnlock(hnew);
	if(clear)
		EmptyClipboard();
	SetClipboardData(format,hnew);
	if(format!=CF_TEXT)
		return;
	HGLOBAL   hloc = GlobalAlloc(GMEM_MOVEABLE, sizeof(LCID));
	LCID *ploc = (LCID *)GlobalLock(hloc);
	*ploc=1049; // LOCALE_USER_DEFAULT
	GlobalUnlock(hloc);
	SetClipboardData(CF_LOCALE,hloc);
}

void SetMemory()
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		if(memory)
		{
			freemem(memory);
			memory=NULL;
		}
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			memory = GetUnique(lpstr);
		}
		GlobalUnlock(hglb); 
	}
	CloseClipboard();
}

void ClearMemory()
{
	if(memory)
	{
		freemem(memory);
		memory=NULL;
	}
}

void GetMemory()
{
	if (!OpenClipboard(hWnd)) 
		return; 
	if(memory)
		SetToClip(memory);
	else
		SetToClip("");
	CloseClipboard();
}

void AddMemory()
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		LPSTR buf=NULL;
		int len=0;
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			buf = GetUnique(lpstr);
			len=strlen(buf);
		}
		GlobalUnlock(hglb); 
		if(memory)
		{
			int memlen=strlen(memory);
			LPSTR newbuf=(LPSTR)malloc(len+memlen+2);
			if(len && buf)
				strcpy(newbuf,buf);
			if(memlen && memory)
			{
				newbuf[len]=' ';
				strcpy(newbuf+len+1,memory);
				LPSTR newbuf2=GetUnique(newbuf);
				SetToClip(newbuf2);
				freemem(newbuf2);
			}
			freemem(newbuf);
		}
		else
			SetToClip(buf);
		freemem(buf);
	} 
	CloseClipboard(); 
}


void DiffMemory()
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		LPSTR buf=NULL;
		int len=0;
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			buf = GetUnique(lpstr);
			if(buf)
				len=strlen(buf);
		}
		GlobalUnlock(hglb); 
		if(buf && memory)
		{
			int memlen=strlen(memory);
			if(memlen && memory)
			{
				int *memarr=(int*)malloc(memlen*sizeof(int));
				memset(memarr,0,memlen*sizeof(int));
				int *bufarr=(int*)malloc(len*sizeof(int));
				memset(bufarr,0,len*sizeof(int));
				int *resarr=(int*)malloc((len+memlen)*sizeof(int));
				int memcount=ParseToIntArr(memory,memarr);
				int bufcount=ParseToIntArr(buf,bufarr);
				int rescount=0;
				for(int j=0; j<memcount; j++)
				{
					bool found=false;
					{
						for(int i=0; i<bufcount; i++)
							if(bufarr[i]==memarr[j])
							{
								found=true;
								break;
							}
					}
					if(!found)
						resarr[rescount++]=memarr[j];
				}
				LPSTR newbuf=ArrToUnique(resarr,rescount);
				freemem(memarr);
				freemem(bufarr);
				freemem(resarr);
				SetToClip(newbuf);
				freemem(newbuf);
			}
			else
				SetToClip(buf);
		}
		else
			SetToClip(buf);
		freemem(buf);
	} 
	CloseClipboard(); 
}

void DupMemory()
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		LPSTR buf=NULL;
		int len=0;
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			buf = GetUnique(lpstr);
			if(buf)
				len=strlen(buf);
		}
		GlobalUnlock(hglb); 
		if(buf && memory)
		{
			int memlen=strlen(memory);
			if(memlen && memory)
			{
				int *memarr=(int*)malloc(memlen*sizeof(int));
				memset(memarr,0,memlen*sizeof(int));
				int *bufarr=(int*)malloc(len*sizeof(int));
				memset(bufarr,0,len*sizeof(int));
				int *resarr=(int*)malloc((len+memlen)*sizeof(int));
				int memcount=ParseToIntArr(memory,memarr);
				int bufcount=ParseToIntArr(buf,bufarr);
				int rescount=0;
				for(int i=0; i<bufcount; i++)
				{
					bool found=false;
					for(int j=0; j<memcount; j++)
					{
						if(bufarr[i]==memarr[j])
						{
							found=true;
							break;
						}
					}
					if(found)
						resarr[rescount++]=bufarr[i];
				}
				LPSTR newbuf=ArrToUnique(resarr,rescount);
				freemem(memarr);
				freemem(bufarr);
				freemem(resarr);
				SetToClip(newbuf);
				freemem(newbuf);
			}
			else
				SetToClip(buf);
		}
		else
			SetToClip(buf);
		freemem(buf);
	} 
	CloseClipboard(); 
}


void MakeSimpleText(void)
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;
	if (!IsClipboardFormatAvailable(CF_TEXT)) 
		if (IsClipboardFormatAvailable(CF_CSV)) 
			getfrm=CF_CSV;
		else
			return;
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		lpstr = (LPSTR)GlobalLock(hglb); 
		int len=0;
		LPSTR buf=NULL;
		if (lpstr != NULL) 
		{ 
			len = strlen(lpstr);
			buf = (LPSTR)malloc(len+1);
			strcpy(buf,lpstr);
		} 
		GlobalUnlock(hglb); 
		SetToClip(buf);
		freemem(buf);
	} 
	CloseClipboard(); 
}


void DoIt(void)
{
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	ClipToUnique(getfrm);
	return; 
}

void MakeCommaList(void)
{
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	ClipToUnique(getfrm);
	return; 
}


void MakeDiesList(void)
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		LPSTR buf=NULL;
		int len=0;
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			buf = GetUnique(lpstr);
			len=strlen(buf);
		}
		GlobalUnlock(hglb); 
		int *arr=(int*)malloc(len*sizeof(int));
		int count=ParseToIntArr(buf,arr);
		freemem(buf);
		LPSTR newbuf=ArrToUnique(arr,count,true);
		freemem(arr);
		SetToClip(newbuf);
		freemem(newbuf);
	} 
	CloseClipboard(); 
}

class CSVRecord
{
public:
	LPSTR *values;
	int count;
	void setValue(int index, LPSTR newvalue) // index=0..
	{
		if(index>=count)
		{
			LPSTR *nv =(LPSTR *)malloc(sizeof(LPSTR *)*(index+1));
			memset(nv,0,sizeof(LPSTR *)*(index+1));
			memcpy(nv,values,sizeof(LPSTR *)*count);
			if(values)
				free(values);
			values=nv;
			count=index+1;
		}
		if(values[index])
			free(values[index]);
		values[index]=(LPSTR)malloc(strlen(newvalue)+1);
		strcpy(values[index],newvalue);
				
	}
	CSVRecord()
	{
		values=NULL;
		count=0;
	}
	~CSVRecord()
	{
		for(int i=0; i<count; i++)
		{
			if(values[i])
				free(values[i]);
		}
	}
};

class CSVSet
{
public:
	CSVRecord **records;
	int rows;
	int cols;
	CSVSet()
	{
		records=NULL;
		rows=0;
		cols=0;
	}
	void clear()
	{
		for(int i =0; i<rows; i++)
			delete records[i];
		rows=0;
		cols=0;
		records=NULL;
	}
	~CSVSet()
	{
		clear();
	}
	void parse(LPSTR text);
	void setValue(int r, int c, LPSTR text) // r=0..., c=0..
	{
		if(r >= rows)
		{
			CSVRecord **nr = (CSVRecord**)malloc(sizeof(CSVRecord*)*(r+1));
			memset(nr,0, sizeof(CSVRecord*)*(r+1));
			memcpy(nr,records, sizeof(CSVRecord*)*rows);
			if(records)
				free(records);
			records=nr;
			rows=r+1;
			records[r]=new CSVRecord();
		}
		records[r]->setValue(c,text);
		if(c>=cols)
			cols=c+1;
	}
	LPSTR value(int r, int c)
	{
		if(r<0 || r>=rows || c<0 || c>=cols)
			return "";
		CSVRecord *v=records[r];
		if(!v || c>=v->count)
			return "";
		LPSTR res = v->values[c];
		if(!res)
			return "";
		return res;
	}
	LPSTR valueByName(int r, LPSTR colname)
	{
		if(r<1 || r>=rows)
			return "";
		CSVRecord *h=records[0];
		if(!h)
			return "";
		int col=-1;
		for(int i=0; i<h->count; i++)
			if(_stricmp(h->values[i],colname)==0)
			{
				col=i;
			}
		CSVRecord *v=records[r];
		if(!v || col<0 || col>=v->count)
			return "";
		return v->values[col];
	}
};

void CSVSet::parse(LPSTR text)
{
	LPSTR p=text;
	const int maxlen=1024;
	CHAR buf[1024];
	LPSTR beg=NULL;
	int row=0;
	while(*p)
	{
		LPSTR k=strchr(p,13);
		if(!k)
		{
			k=strchr(p,10);
			if(!k)
				k=p+strlen(p);
		}
		int col=0;
		while(*p && p<k)
		{
			if(*p != '"')
			{
				beg=p;
				LPSTR n=strchr(p,',');
				if(!n)
					n=p+strlen(p);
				if(n>k)
					n=k;
				p=n;
				int len=p-beg;
				if(len>=maxlen)
					len=maxlen-1;
				strncpy(buf,beg,len);
				buf[len]=0;
				setValue(row,col++,buf);
				if(*p==',')
					p++;
			}
			else
			{
				p++;
				beg=p;
				LPSTR dest=buf;
				int len=0;
				bool is_end=false;
				while(!is_end && len<maxlen-1)
				{
					CHAR c=*p++;
					switch(c)
					{
					case '"':
						if(*p=='"')
						{
							*dest++='"';
							len++;
							p++;
						}
						else
						{
							*dest++=0;
							len++;
							is_end=true;
						}
						break;
					//case ',':
					//	*dest++=0;
					//	len++;
					//	is_end=true;
					//	break;
					default:
						*dest++=c;
						len++;
					}
				}
				buf[len]=0;
				setValue(row,col++,buf);
				if(*p==',')
					p++;
			}
		}
		if(*p==13)
			p++;
		if(*p==10)
			p++;
		row++;
	}
}



LPSTR *addline(LPSTR *dest, LPSTR src)
{
	int newlen= strlen(*dest)+ strlen(src);
	LPSTR res=(LPSTR)malloc(newlen+1);
	strcpy(res,*dest);
	strcat(res,src);
	free(*dest);
	*dest=res;
	return dest;
}



void SetToClipHTML(LPSTR text)
{
	const LPSTR header=
		"Version:0.9\r\n"
		"StartHTML:-1\r\n"
		"EndHTML:-1\r\n"
		"StartFragment:000081\r\n"
		"EndFragment:";
	int widelen=strlen(text)+1;
	int widesize=widelen*sizeof(WCHAR);
	LPWSTR wide=(LPWSTR)malloc(widesize);
	MultiByteToWideChar(1251,0,text,-1,wide,widelen);
	int utflen=widelen*4+1;
	LPSTR utf=(LPSTR)malloc(utflen);
	int utfcount = WideCharToMultiByte(CP_UTF8,0,wide,-1,utf,utflen,NULL,NULL);
	//int utfcount = UnicodeToUtf8(wide,utf,utflen);
	utf[utfcount]=0;
	free(wide);

	int hlen=strlen(header);
	char end[16];
	_itoa(utfcount+hlen+2+6+100000,end,10);
	end[0]--;
	LPSTR res=(LPSTR)malloc(1);
	res[0]=0;
	addline(&res,header);
	addline(&res,end);
	addline(&res,"\r\n");
	addline(&res,utf);
	//SetToClip(res);
	SetToClip(res,CF_HTML);
	free(res);
}

/*
void CSVToTable(LPSTR text, LPSTR *list, LPSTR *html)
{
	CSVSet set;
	set.parse(text);
	*html=(LPSTR)malloc(1);
	*html[0]=0;
	addline(html,"<html>\r\n<body>\r\n<table border=1 cellpadding=3 cellspacing=0 bordercolor=#000000 style='border-collapse:collapse;border:none'>\r\n");
	*list=(LPSTR)malloc(1);
	*list[0]=0;
	for(int r=1; r<set.rows; r++)
	{
		addline(html,"<tr>");
		LPSTR id = set.valueByName(r,"Id");
		LPSTR title = set.valueByName(r,"Title");
		if((id && *id) || (title && *title))
		{
			addline(html,"<td valign='top'>");
			addline(html,id);
			addline(html,"</td>");
			addline(html,"<td valign='top'>");
			addline(html,title);
			addline(html,"</td>");
			addline(list,id);
			addline(list,"\t");
			addline(list,title);
			addline(list,"\r\n");
		}
		addline(html,"</tr>\r\n");
	}
	addline(html,"</table>\r\n</body>\r\n</html>\r\n");
}


*/

void MakeTableIdTitle()
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		LPSTR list=NULL;
		LPSTR html=NULL;
		int len=0;
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			//CSVToTable(lpstr,&list,&html);
			CSVSet set;
			set.parse(lpstr);
			html=(LPSTR)malloc(1);
			html[0]=0;
			addline(&html,"<html>\r\n<body>\r\n<table border=1 cellpadding=3 cellspacing=0 bordercolor=#000000 style='border-collapse:collapse;border:none'>\r\n");
			list=(LPSTR)malloc(1);
			list[0]=0;
			for(int r=1; r<set.rows; r++)
			{
				addline(&html,"<tr>");
				LPSTR id = set.valueByName(r,"Id");
				LPSTR title = set.valueByName(r,"Title");
				if((id && *id) || (title && *title))
				{
					addline(&html,"<td valign='top'>");
					addline(&html,id);
					addline(&html,"</td>");
					addline(&html,"<td valign='top'>");
					addline(&html,title);
					addline(&html,"</td>");
					addline(&list,id);
					addline(&list,"\t");
					addline(&list,title);
					addline(&list,"\r\n");
				}
				addline(&html,"</tr>\r\n");
			}
			addline(&html,"</table>\r\n</body>\r\n</html>\r\n");
		}
		GlobalUnlock(hglb); 
		SetToClipHTML(html);
		SetToClip(list,CF_TEXT,false);
		freemem(list);
		freemem(html);
	} 
	CloseClipboard(); 
}

void MakeTableFull()
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		LPSTR list=NULL;
		LPSTR html=NULL;
		int len=0;
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			//CSVToTable(lpstr,&list,&html);
			CSVSet set;
			set.parse(lpstr);
			html=(LPSTR)malloc(1);
			html[0]=0;
			addline(&html,"<html>\r\n<body>\r\n<table border=1 cellpadding=3 cellspacing=0 bordercolor=#000000 style='border-collapse:collapse;border:none'>\r\n");
			list=(LPSTR)malloc(1);
			list[0]=0;
			for(int r=0; r<set.rows; r++)
			{
				addline(&html,"<tr>");
				for(int c=0; c<set.cols; c++)
				{
					LPSTR v=set.value(r,c);
					addline(&html,"<td valign='top'>");
					addline(&html,v);
					addline(&html,"</td>");
					if(c)
						addline(&list, "\t");
					addline(&list, v);
				}
				addline(&html,"</tr>\r\n");
				addline(&list, "\r\n");
			}
			addline(&html,"</table>\r\n</body>\r\n</html>\r\n");
		}
		GlobalUnlock(hglb); 
		SetToClipHTML(html);
		SetToClip(list,CF_TEXT,false);
		freemem(list);
		freemem(html);
	} 
	CloseClipboard(); 
}

void DiffMemoryTable()
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		LPSTR buf=NULL;
		int len=0;
		lpstr = (LPSTR)GlobalLock(hglb); 
		if (lpstr != NULL) 
		{ 
			buf = GetUnique(lpstr);
			if(buf)
				len=strlen(buf);
		}
		GlobalUnlock(hglb); 
		int memcount=0;
		int bufcount=0;
		int *bufarr=NULL;
		int *memarr=NULL;
		if(len)
		{
			bufarr=(int*)malloc(len*sizeof(int));
			memset(bufarr,0,len*sizeof(int));
		}
		if(buf && bufarr)
			bufcount=ParseToIntArr(buf,bufarr);
		int memlen=0;
		if(memory)
			memlen=strlen(memory);
		if(memlen && memory)
		{
			memarr=(int*)malloc(memlen*sizeof(int));
			memset(memarr,0,memlen*sizeof(int));
			memcount=ParseToIntArr(memory,memarr);
		}
		//int rescount=0;
		int i=0, j=0;
		char val[20];
		const LPSTR td="<td valign='top'>";
		const LPSTR ed="</td>";
		const LPSTR crlf="\r\n";
		const LPSTR empty="<td valign='top'></td>";
		LPSTR html=(LPSTR)malloc(1);
		html[0]=0;
		addline(&html,"<html>\r\n<body>\r\n<table border=1 cellpadding=3 cellspacing=0 bordercolor=#000000 style='border-collapse:collapse;border:none'>\r\n");
		LPSTR list=(LPSTR)malloc(1);
		list[0]=0;
		addline(&html,
			"<tr>"
			"<td valign='top'>Память</td>"
			"<td valign='top'>Буфер</td>"
			"<td valign='top'>Совпадающие</td>"
			"<td valign='top'>Нет в памяти</td>"
			"<td valign='top'>Нет в буфере</td>"
			"</tr>\r\n"
			);
		while(i<bufcount || j<memcount)
		{
			addline(&html,"<tr>");
			if(i<bufcount)
			{
				if(j<memcount)
				{
					if(bufarr[i]<memarr[j])
					{
						_itoa(bufarr[i++],val,10);
						addline(&html,empty);
						addline(&html,td);	addline(&html,val);	addline(&html,ed);
						addline(&html,empty);
						addline(&html,td);	addline(&html,val);	addline(&html,ed);
						addline(&html,empty);
						addline(&list,"\t");
						addline(&list,val);
						addline(&list,crlf);
					}
					else if(bufarr[i]>memarr[j])
					{
						_itoa(memarr[j++],val,10);
						addline(&html,td); addline(&html,val); addline(&html,ed);
						addline(&html,empty);
						addline(&html,empty);
						addline(&html,empty);
						addline(&html,td); addline(&html,val); addline(&html,ed);
						addline(&list,val);
						addline(&list,"\t");
						addline(&list,"\r\n");
					}
					else
					{
						_itoa(bufarr[i++],val,10);
						j++;
						addline(&html,td); addline(&html,val); addline(&html,ed);
						addline(&html,td); addline(&html,val); addline(&html,ed);
						addline(&html,td); addline(&html,val); addline(&html,ed);
						addline(&html,empty);
						addline(&html,empty);
						addline(&list,val);
						addline(&list,"\t");
						addline(&list,val);
						addline(&list,"\r\n");
					}
				}
				else
				{
					_itoa(bufarr[i++],val,10);
					addline(&html,empty);
					addline(&html,td); addline(&html,val); addline(&html,ed);
					addline(&html,empty);
					addline(&html,td); addline(&html,val); addline(&html,ed);
					addline(&html,empty);
					addline(&list,"\t");
					addline(&list,val);
					addline(&list,"\r\n");
				}
			}
			else
			{
				_itoa(memarr[j++],val,10);
				addline(&html,td); addline(&html,val); addline(&html,ed);
				addline(&html,empty);
				addline(&html,empty);
				addline(&html,empty);
				addline(&html,td); addline(&html,val); addline(&html,ed);
				addline(&list,val);
				addline(&list,"\t");
				addline(&list,"\r\n");
			}
			addline(&html,"</tr>");
		}
		addline(&html,"</table>\r\n</body>\r\n</html>\r\n");
		freemem(memarr);
		SetToClipHTML(html);
		SetToClip(list,CF_TEXT,false);
		freemem(html);
		freemem(list);
		freemem(bufarr);
		freemem(buf);
	} 
	CloseClipboard(); 
}



void MakeColumnList(void)
{
	HGLOBAL   hglb; 
	LPSTR    lpstr; 
	UINT getfrm=CF_TEXT;

	if (!IsClipboardFormatAvailable(CF_TEXT)) 
	{
		if (!IsClipboardFormatAvailable(CF_CSV)) 
			return;
		getfrm=CF_CSV;
	}
	if (!OpenClipboard(hWnd)) 
		return; 
	hglb = GetClipboardData(getfrm); 
	if (hglb != NULL) 
	{ 
		LPSTR buf=NULL;
		int len=0;
		lpstr = (LPSTR)GlobalLock(hglb); 
		len=strlen(lpstr);
		int *arr=(int*)malloc(len*sizeof(int));
		int count=0;
		if (lpstr != NULL) 
		{ 
			count=ParseToIntArr(lpstr,arr);
			qsort(arr,count,sizeof(arr[0]),compare);
		}
		GlobalUnlock(hglb); 
		LPSTR html =(LPSTR)malloc(1);
		html[0]=0;
		addline(&html,"<html>\r\n<body>\r\n<table border=1 cellpadding=3 cellspacing=0 bordercolor=#000000 style='border-collapse:collapse;border:none'>\r\n");
		LPSTR list =(LPSTR)malloc(1);
		list[0]=0;
		int prev=0;
		bool first=true;
		for(int i=0; i<count; i++)
		{
			if(CHECK_MIN && arr[i] < MIN_NUMBER)
				continue;
			if(arr[i] == prev)
				continue;
			addline(&html,"<tr><td valign='top'>");
			char dest[20];
			_itoa(arr[i],dest,10);
			addline(&html,dest);
			addline(&html,"</td></tr>\r\n");
			addline(&list,dest);
			addline(&list,"\r\n");
			prev=arr[i];
		}
		addline(&html,"</table>\r\n</body>\r\n</html>\r\n");
		SetToClipHTML(html);
		SetToClip(list,CF_TEXT,false);
		freemem(html);
		freemem(list);
	}
	CloseClipboard(); 
}




//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	switch (message)
	{
	case ICON_NOTIFY:
		switch (lParam)
		{
		case WM_LBUTTONDOWN:
			DoIt();
			break;
		case WM_RBUTTONUP:
			//DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			POINT pt;
			GetCursorPos(&pt);
            //int x = GET_X_LPARAM(lParam); 
            //int y = GET_Y_LPARAM(lParam); 
			HMENU menu;
			HMENU submenu;
			menu = LoadMenu(hInst,MAKEINTRESOURCE(IDC_POPUP));
			submenu = GetSubMenu(menu,0);
			SetForegroundWindow(hWnd);
			TrackPopupMenu(submenu,
                     TPM_RIGHTBUTTON,
                     pt.x,
                     pt.y,
                     0,
                     hWnd,
                     NULL);
			PostMessage(hWnd, WM_NULL, 0, 0);
			DestroyMenu(menu);
			break;
		}
		break;
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		// Parse the menu selections:
		switch (wmId)
		{
		case IDM_COMMA:
			MakeCommaList();
			break;
		case IDM_STORE:
			SetMemory();
			break;
		case IDM_ADD:
			AddMemory();
			break;
		case IDM_DIFF:
			DiffMemory();
			break;
		case IDM_DUP:
			DupMemory();
			break;
		case IDM_GET:
			GetMemory();
			break;
		case IDM_DIES:
			MakeDiesList();
			break;
		case IDM_CLEAR:
			ClearMemory();
			break;
		case IDM_LIST:
			MakeTableIdTitle();
			break;
		case IDM_HTML:
			MakeTableFull();
			break;
		case IDM_TEXT:
			MakeSimpleText();
			break;
		case IDM_CMP:
			DiffMemoryTable();
			break;
		case IDM_COLUMN:
			MakeColumnList();
			break;
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			RemoveIcon(hWnd);
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: Add any drawing code here...
		EndPaint(hWnd, &ps);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}
