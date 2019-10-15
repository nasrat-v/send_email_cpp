
#include <tchar.h>
#include <Windows.h>

#include "EASendMailObj.tlh"
using namespace EASendMailObjLib;

const int ConnectNormal = 0;
const int ConnectSSLAuto = 1;
const int ConnectSTARTTLS = 2;
const int ConnectDirectSSL = 3;
const int ConnectTryTLS = 4;

int _tmain(int argc, _TCHAR* argv[])
{
	::CoInitialize(NULL);

	IMailPtr oSmtp = NULL;
	oSmtp.CreateInstance(__uuidof(EASendMailObjLib::Mail));
	oSmtp->LicenseCode = _T("TryIt");

	// Set your sender email address
	oSmtp->FromAddr = _T("");
	// Add recipient email address
	oSmtp->AddRecipientEx(_T(""), 0);

	// Set email subject
	oSmtp->Subject = _T("simple email from Visual C++ project");
	// Set email body
	oSmtp->BodyText = _T("this is a test email sent from Visual C++ project, do not reply");

	// Your SMTP server address
	oSmtp->ServerAddr = _T("smtp.office365.com");

	// User and password for ESMTP authentication, if your server doesn't
	// require User authentication, please remove the following codes.
	oSmtp->UserName = _T("");
	oSmtp->Password = _T("");

	// Most mordern SMTP servers require SSL/TLS connection now.
	// ConnectTryTLS means if server supports SSL/TLS, SSL/TLS will be used automatically.
	oSmtp->ConnectType = ConnectTryTLS;

	// If your SMTP server uses 587 port
	// oSmtp->ServerPort = 587;

	// If your SMTP server requires SSL/TLS connection on 25/587/465 port
	// oSmtp->ServerPort = 25; // 25 or 587 or 465
	// oSmtp->ConnectType = ConnectSSLAuto;

	_tprintf(_T("Start to send ...\r\n"));

	if (oSmtp->SendMail() == 0)
	{
		_tprintf(_T("sent successfully!\r\n"));
	}
	else
	{
		_tprintf(_T("failed to send with the following error: %s\r\n"),
			(const TCHAR*)oSmtp->GetLastErrDescription());
	}

	return 0;
}