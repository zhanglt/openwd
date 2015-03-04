/// ////////////////////////////////////////////////////////////////////////////////
/// 版权所有：Copyright (C) Copyright 2009
/// 模块名称：des.cpp
/// 模块编号：0
/// 文件名称：d:\wps\OperationDLL\des.cpp
/// 功    能：正文传输过程中的密钥处理 
/// 作    者：zhanglt
/// 创建时间：2009-10-24 20:15:28
/// 修改时间：
/// 公    司： 
/// 产    品：WPSOCX 
/// ////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PubFunction.h"
#include "des.h"
#include <comutil.h>
#include <stdio.h>
#include <comdef.h>
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "kernel32.lib")

char SubKeys[16][48];//储存16组48位密钥
char szCiphertext[16];//储存16位密文(十六进制字符串)
char szPlaintext[8];//储存8位明文字符串
char szFCiphertextAnyLength[8192];//任意长度密文(十六进制字符串)
char szFPlaintextAnyLength[4096];//任意长度明文字符串


void CreateSubKey(char* sz)
{
	char szTmpL[28] = { 0 };
	char szTmpR[28] = { 0 };
	char szCi[28] = { 0 };
	char szDi[28] = { 0 };
	memcpy(szTmpL, sz, 28);
	memcpy(szTmpR, sz + 28, 28);

	for (int i = 0; i<16; i++)
	{
		//shift to left
		//Left 28 bits
		memcpy(szCi, szTmpL + Shift_Table[i], 28 - Shift_Table[i]);
		memcpy(szCi + 28 - Shift_Table[i], szTmpL, Shift_Table[i]);
		//Right 28 bits
		memcpy(szDi, szTmpR + Shift_Table[i], 28 - Shift_Table[i]);
		memcpy(szDi + 28 - Shift_Table[i], szTmpR, Shift_Table[i]);

		//permuted choice 48 bits key
		char szTmp56[56] = { 0 };
		memcpy(szTmp56, szCi, 28);
		memcpy(szTmp56 + 28, szDi, 28);
		for (int j = 0; j<48; j++)
		{
			SubKeys[i][j] = szTmp56[PC2_Table[j] - 1];
		}
		//Evaluate new szTmpL and szTmpR
		memcpy(szTmpL, szCi, 28);
		memcpy(szTmpR, szDi, 28);
	}
}

string HexCharToBinary(char ch)
{
	switch (ch)
	{
	case '0':
		return "0000";
	case '1':
		return "0001";
	case '2':
		return "0010";
	case '3':
		return "0011";
	case '4':
		return "0100";
	case '5':
		return "0101";
	case '6':
		return "0110";
	case '7':
		return "0111";
	case '8':
		return "1000";
	case '9':
		return "1001";
	case 'a':
		return "1010";
	case 'b':
		return "1011";
	case 'c':
		return "1100";
	case 'd':
		return "1101";
	case 'e':
		return "1110";
	case 'f':
		return "1111";
	default:
		return "";
	}
}



char* HexIntToBinary(int i)
{
	switch (i)
	{
	case 0:
		return "0000";
	case 1:
		return "0001";
	case 2:
		return "0010";
	case 3:
		return "0011";
	case 4:
		return "0100";
	case 5:
		return "0101";
	case 6:
		return "0110";
	case 7:
		return "0111";
	case 8:
		return "1000";
	case 9:
		return "1001";
	case 10:
		return "1010";
	case 11:
		return "1011";
	case 12:
		return "1100";
	case 13:
		return "1101";
	case 14:
		return "1110";
	case 15:
		return "1111";
	default:
		return "";
	}
}


string BinaryToString(char* szSource, int len, bool bType)
{
	//bType == true is Binary to Hex
	//else is Binary to Char
	int ilen;
	if (len % 8 != 0)
		return "";
	else
		ilen = len / 8;

	string s_return = "";
	for (int i = 0; i<ilen; i++)
	{
		char szTmp8[8] = { 0 };
		double iCh = 0;
		memcpy(szTmp8, szSource + 8 * i, 8);
		for (int j = 0; j<8; j++)
		{
			iCh += SingleCharToBinary(szTmp8[j]) * pow(2.0, 7 - j);
		}
		if (bType)
		{
			char buffer[2] = { 0 };
			_itoa_s(iCh, buffer, 16);
			//if the integer less than 16,insert a zero to buffer
			if (iCh < 16)
			{
				char cTmp = buffer[0];
				buffer[0] = '0';
				buffer[1] = cTmp;
			}
			s_return += buffer[0];
			s_return += buffer[1];
			buffer[0] = '\0';
			buffer[1] = '\0';
		}
		else
		{
			char ch = (char)iCh;
			s_return += ch;
		}
	}

	return s_return;
}

void EncryptData(string s)
{
	char sz_IP[64] = { 0 };
	char sz_Li[32] = { 0 };
	char sz_Ri[32] = { 0 };
	char sz_Final64[64] = { 0 };
	char szCiphertextBinary[64] = { 0 };
	//IP
	InitialPermuteData(s, sz_IP, true);
	memcpy(sz_Li, sz_IP, 32);
	memcpy(sz_Ri, sz_IP + 32, 32);

	for (int i = 0; i<16; i++)
	{
		FunctionF(sz_Li, sz_Ri, i);
	}
	//so D=LR
	memcpy(sz_Final64, sz_Li, 32);
	memcpy(sz_Final64 + 32, sz_Ri, 32);

	//~IP
	for (int j = 0; j<64; j++)
	{
		szCiphertextBinary[j] = sz_Final64[IPR_Table[j] - 1];
	}
	memcpy(szCiphertext, BinaryToString(szCiphertextBinary, 64, true).c_str(), 16);
}

void DecryptData(string s)
{
	char sz_IP[64] = { 0 };
	char sz_Li[32] = { 0 };
	char sz_Ri[32] = { 0 };
	char sz_Final64[64] = { 0 };
	char szPlaintextBinary[64] = { 0 };
	//IP --- return is sz_IP
	InitialPermuteData(s, sz_IP, false);
	//divide the 64 bits data to two parts
	memcpy(sz_Ri, sz_IP, 32); //exchange L to R
	memcpy(sz_Li, sz_IP + 32, 32);  //exchange R to L

	//16 rounds F and xor and exchange
	for (int i = 0; i<16; i++)
	{
		FunctionF(sz_Li, sz_Ri, 15 - i);
	}
	//the round 16 will not exchange L and R
	//so D=LR is D=RL
	memcpy(sz_Final64, sz_Ri, 32);
	memcpy(sz_Final64 + 32, sz_Li, 32);

	// ~IP
	for (int j = 0; j<64; j++)
	{
		szPlaintextBinary[j] = sz_Final64[IPR_Table[j] - 1];
	}
	memcpy(szPlaintext, BinaryToString(szPlaintextBinary, 64, false).c_str(), 8);
}

void FunctionF(char* sz_Li, char* sz_Ri, int iKey)
{
	char sz_48R[48] = { 0 };
	char sz_xor48[48] = { 0 };
	char sz_P32[32] = { 0 };
	char sz_Rii[32] = { 0 };
	char sz_Key[48] = { 0 };
	memcpy(sz_Key, SubKeys[iKey], 48);
	ExpansionR(sz_Ri, sz_48R);
	XOR(sz_48R, sz_Key, 48, sz_xor48);
	string s_Compress32 = CompressFuncS(sz_xor48);
	PermutationP(s_Compress32, sz_P32);
	XOR(sz_P32, sz_Li, 32, sz_Rii);
	memcpy(sz_Li, sz_Ri, 32);
	memcpy(sz_Ri, sz_Rii, 32);
}

void InitialPermuteData(string s, char* Return_value, bool bType)
{
	//if bType==true is encrypt
	//else is decrypt
	if (bType)
	{
		char sz_64data[64] = { 0 };
		int iTmpBit[64] = { 0 };
		for (int i = 0; i<64; i++)
		{
			iTmpBit[i] = (s[i >> 3] >> (i & 7)) & 1;
			//a = 0x61 = 0110,0001
			//after this , a is 1000,0110

		}
		//let me convert it to right
		for (int j = 0; j<8; j++)
		for (int k = 0; k<8; k++)
			sz_64data[8 * j + k] = SingleBinaryToChar(iTmpBit[8 * (j + 1) - (k + 1)]);
		//IP
		char sz_IPData[64] = { 0 };
		for (int k = 0; k<64; k++)
		{
			sz_IPData[k] = sz_64data[IP_Table[k] - 1];
		}
		memcpy(Return_value, sz_IPData, 64);
	}
	else
	{
		string sz_64data;
		for (unsigned int ui = 0; ui<s.length(); ui++)
		{
			char ch = s[ui];
			sz_64data += HexCharToBinary(tolower(ch));
		}
		char sz_IPData[64] = { 0 };
		for (int i = 0; i<64; i++)
		{
			sz_IPData[i] = sz_64data[IP_Table[i] - 1];
		}
		memcpy(Return_value, sz_IPData, 64);
	}

}


void ExpansionR(char* sz, char* Return_value)
{
	char sz_48ER[48] = { 0 };
	for (int i = 0; i<48; i++)
	{
		sz_48ER[i] = sz[E_Table[i] - 1];
	}
	memcpy(Return_value, sz_48ER, 48);
}


void XOR(char* sz_P1, char* sz_P2, int len, char* Return_value)
{
	char sz_Buffer[256] = { 0 };
	for (int i = 0; i<len; i++)
	{
		sz_Buffer[i] = SingleBinaryToChar(SingleCharToBinary(sz_P1[i]) ^ SingleCharToBinary(sz_P2[i]));
	}
	memcpy(Return_value, sz_Buffer, len);
}

string CompressFuncS(char* sz_48)
{
	char sTmp[8][6] = { 0 };
	string sz_32 = "";
	for (int i = 0; i<8; i++)
	{
		memcpy(sTmp[i], sz_48 + 6 * i, 6);
		int iX = SingleCharToBinary(sTmp[i][0]) * 2 + SingleCharToBinary(sTmp[i][5]);
		int iY = 0;
		for (int j = 1; j<5; j++)
		{
			iY += SingleCharToBinary(sTmp[i][j]) * pow(2.0, 4 - j);
		}
		sz_32 += HexIntToBinary(S_Box[i][iX][iY]);
	}
	return sz_32;
}


void PermutationP(string s, char* Return_value)
{
	char sz_32bits[32] = { 0 };
	for (int i = 0; i<32; i++)
	{
		sz_32bits[i] = s[P_Table[i] - 1];
	}
	memcpy(Return_value, sz_32bits, 32);
}

int SingleCharToBinary(char ch)
{
	if (ch == '1')
		return 1;
	else
		return 0;
}

char SingleBinaryToChar(int i)
{
	if (i == 1)
		return '1';
	else
		return '0';
}

void SetCiphertext(char* value)
{
	memcpy(szCiphertext, value, 16);
}
char* GetCiphertext()
{
	return szCiphertext;
}



void SetPlaintext(char* value)
{
	memcpy(szPlaintext, value, 8);
}
char* GetPlaintext()
{
	return szPlaintext;
}


//fill the data to 8 bits
string FillToEightBits(string sz)
{
	//length less than 8 , add zero(s) to tail
	switch (sz.length())
	{
	case 7:
		sz += "$";
		break;
	case 6:
		sz += "$$";
		break;
	case 5:
		sz += "$$$";
		break;
	case 4:
		sz += "$$$$";
		break;
	case 3:
		sz += "$$$$$";
		break;
	case 2:
		sz += "$$$$$$";
		break;
	case 1:
		sz += "$$$$$$$";
		break;
	default:
		break;
	}
	return sz;
}





void CleanPlaintextMark(int iPlaintextLength)
{
	char szLast7Chars[7] = { 0 };
	memcpy(szLast7Chars, szFPlaintextAnyLength + iPlaintextLength - 7, 7);
	for (int i = 0; i<7; i++)
	{
		char* pDest = strrchr(szLast7Chars, '$');
		if (pDest == NULL)
		{
			break;
		}
		else
		{
			int iPos = (int)(pDest - szLast7Chars + 1);
			if (iPos != 7 - i)
			{
				break;
			}
			else
			{
				szLast7Chars[6 - i] = '\0';
				continue;
			}
		}
	}
	memcpy(szFPlaintextAnyLength + iPlaintextLength - 7, szLast7Chars, 7);
}


void InitializeKey(string s)
{

	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		USES_CONVERSION;
	//convert 8 char-bytes key to 64 binary-bits
	char sz_64key[64] = { 0 };
	int iTmpBit[64] = { 0 };
	for (int i = 0; i<64; i++)
	{
		iTmpBit[i] = (s[i >> 3] >> (i & 7)) & 1;
		//a = 0x61 = 0110,0001
		//after this , a is 1000,0110

	}
	//let me convert it to right
	for (int j = 0; j<8; j++)
	for (int k = 0; k<8; k++)
		sz_64key[8 * j + k] = SingleBinaryToChar(iTmpBit[8 * (j + 1) - (k + 1)]);
	//PC 1
	char sz_56key[56] = { 0 };
	for (int k = 0; k<56; k++)
	{
		sz_56key[k] = sz_64key[PC1_Table[k] - 1];
	}
	CreateSubKey(sz_56key);

	//	return S_OK;
}



CString  DecryptAnyLength1(BSTR szSource)

{
	string   szstr;

	szstr = (LPCTSTR)(_bstr_t)szSource;


	int iLength = szstr.length();
	int iRealLengthOfPlaintext = 0;
	//if the length is 16 , call DecyptData
	if (iLength == 16)
	{
		DecryptData(szstr);
		memcpy(szFPlaintextAnyLength, szPlaintext, 8);
		iRealLengthOfPlaintext = 8;
	}
	//it's not impossible the length less than 16
	else if (iLength < 16)
	{
		sprintf(szFPlaintextAnyLength, "Please enter your correct cipertext!");
	}
	//else if then lenth bigger than 16
	//divide the string to multi-parts
	else if (iLength > 16)
	{
		double iParts = ceil(iLength / 16.0);
		double iResidue = iLength % 16;
		//if the data can't be divided exactly by 16
		//it's meaning the data is a wrong !
		if (iResidue != 0)
		{
			sprintf(szFPlaintextAnyLength, "Please enter your correct cipertext!");
			return  "error";;
		}
		iRealLengthOfPlaintext = iParts * 8;
		//Decrypt the data 16 by 16
		for (int i = 0; i<iParts; i++)
		{
			string szTmp = szstr.substr(i * 16, 16);
			DecryptData(szTmp);
			//after call DecryptData
			//cpoy the temp result to szFPlaintextAnyLength
			memcpy(szFPlaintextAnyLength + 8 * i, szPlaintext, 8);
		}
	}
	//find and clear the mark
	//which is add by program when ciphertext is less than 8
	CleanPlaintextMark(iRealLengthOfPlaintext);

	//*pval=_com_util::ConvertStringToBSTR(szFPlaintextAnyLength);
	//SysFreeString(szFCiphertextAnyLength);
	return szFPlaintextAnyLength;
}

BOOL GetUnlokPassword(char * Password)
{

	CString szValue;
	//读取注册表中的password的密文。
	szValue = AfxGetApp()->GetProfileString("openwd", "Password", "");

	//    MessageBox(NULL,szValue,"注册表中取出的password",MB_OK|MB_ICONINFORMATION);
	//密文是否为空
	if (szValue == "")
	{
		//如果为空赋值为初始密码
		szValue = "openwdoa";
	}
	else{
		//对密文进行解密操作
		InitializeKey("12345678");
		//MessageBox(NULL,szValue,"开始解码后的password",MB_OK|MB_ICONINFORMATION);
		szValue = DecryptAnyLength1(_com_util::ConvertStringToBSTR(szValue));
		//MessageBox(NULL,szValue,"解码后的password",MB_OK|MB_ICONINFORMATION);
	}

	strcpy(Password, szValue);
	return true;

}