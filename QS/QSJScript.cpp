// QSJScript.cpp : Implementation of CQSJScript

#include "stdafx.h"
#include "QSJScript.h"
#include "QSScriptSite.h"
#include "..\QSUtil\OutputDebugVariant.h"
#include "..\QSUtil\VariantToBSTR.h"

// CQSJScript

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSJScript::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSJScript
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

HRESULT CQSJScript::FinalConstruct()
{
	HRESULT hr = S_OK;

	CHECKHR(CQSScriptSite::_CreatorClass::CreateInstance(NULL, IID_IQSScriptSite, (void**) &m_spIQSScriptSite));
	CHECKHR(m_spIQSScriptSite->put_ScriptEngine(CComBSTR(L"JScript")));

	CComBSTR bstrScript(
		L"function jsEncodeURIComponent(uri) { return encodeURIComponent(uri); }\n"
		L"function jsTest(x,y,z) { return x + x; }\n"
		L"function jsParse(s) { try { return JSON.parse(s); } catch (err) { return err; } }\n"
		L"function jsGetObjKey(o,k) { return o[k]; }\n"
		L"function jsSetObjKey(o,k,v) { o[k] = v; }\n"
		L"function jsStringify(o,r,s) { return JSON.stringify(o,r,s); }\n"
		);
	CHECKHR(m_spIQSScriptSite->Execute(bstrScript, CComVariant()));

	CHECKHR(m_spIQSScriptSite->Execute(CComBSTR(L"\"object\"!=typeof JSON&&(JSON={}),function(){\"use strict\";function f(t){return 10>t?\"0\"+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=0,rx_escapable.test(t)?'\"'+t.replace(rx_escapable,function(t){var e=meta[t];return\"string\"==typeof e?e:\"\\\\u\"+(\"0000\"+t.charCodeAt(0).toString(16)).slice(-4)})+'\"':'\"'+t+'\"'}function str(t,e){var r,n,o,u,f,a=gap,i=e[t];switch(i&&\"object\"==typeof i&&\"function\"==typeof i.toJSON&&(i=i.toJSON(t)),\"function\"==typeof rep&&(i=rep.call(e,t,i)),typeof i){case\"string\":return quote(i);case\"number\":return isFinite(i)?i+\"\":\"null\";case\"boolean\":case\"null\":return i+\"\";case\"object\":if(!i)return\"null\";if(gap+=indent,f=[],\"[object Array]\"===Object.prototype.toString.apply(i)){for(u=i.length,r=0;u>r;r+=1)f[r]=str(r,i)||\"null\";return o=0===f.length?\"[]\":gap?\"[\\n\"+gap+f.join(\",\\n\"+gap)+\"\\n\"+a+\"]\":\"[\"+f.join(\",\")+\"]\",gap=a,o}if(rep&&\"object\"==typeof rep)for(u=rep.length,r=0;u>r;r+=1)\"string\"==typeof rep[r]&&(n=rep[r],o=str(n,i),o&&f.push(quote(n)+(gap?\": \":\":\")+o));else for(n in i)Object.prototype.hasOwnProperty.call(i,n)&&(o=str(n,i),o&&f.push(quote(n)+(gap?\": \":\":\")+o));return o=0===f.length?\"{}\":gap?\"{\\n\"+gap+f.join(\",\\n\"+gap)+\"\\n\"+a+\"}\":\"{\"+f.join(\",\")+\"}\",gap=a,o}}var rx_one=/^[\\],:{}\\s]*$/,rx_two=/\\\\(?:[\"\\\\\\/bfnrt]|u[0-9a-fA-F]{4})/g,rx_three=/\"[^\"\\\\\\n\\r]*\"|true|false|null|-?\\d+(?:\\.\\d*)?(?:[eE][+\\-]?\\d+)?/g,rx_four=/(?:^|:|,)(?:\\s*\\[)+/g,rx_escapable=/[\\\\\\\"\\u0000-\\u001f\\u007f-\\u009f\\u00ad\\u0600-\\u0604\\u070f\\u17b4\\u17b5\\u200c-\\u200f\\u2028-\\u202f\\u2060-\\u206f\\ufeff\\ufff0-\\uffff]/g,rx_dangerous=/[\\u0000\\u00ad\\u0600-\\u0604\\u070f\\u17b4\\u17b5\\u200c-\\u200f\\u2028-\\u202f\\u2060-\\u206f\\ufeff\\ufff0-\\uffff]/g;\"function\"!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+\"-\"+f(this.getUTCMonth()+1)+\"-\"+f(this.getUTCDate())+\"T\"+f(this.getUTCHours())+\":\"+f(this.getUTCMinutes())+\":\"+f(this.getUTCSeconds())+\"Z\":null},Boolean.prototype.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value);var gap,indent,meta,rep;\"function\"!=typeof JSON.stringify&&(meta={\"\\b\":\"\\\\b\",\"	\":\"\\\\t\",\"\\n\":\"\\\\n\",\"\\f\":\"\\\\f\",\"\\r\":\"\\\\r\",'\"':'\\\\\"',\"\\\\\":\"\\\\\\\\\"},JSON.stringify=function(t,e,r){var n;if(gap=\"\",indent=\"\",\"number\"==typeof r)for(n=0;r>n;n+=1)indent+=\" \";else\"string\"==typeof r&&(indent=r);if(rep=e,e&&\"function\"!=typeof e&&(\"object\"!=typeof e||\"number\"!=typeof e.length))throw Error(\"JSON.stringify\");return str(\"\",{\"\":t})}),\"function\"!=typeof JSON.parse&&(JSON.parse=function(text,reviver){function walk(t,e){var r,n,o=t[e];if(o&&\"object\"==typeof o)for(r in o)Object.prototype.hasOwnProperty.call(o,r)&&(n=walk(o,r),void 0!==n?o[r]=n:delete o[r]);return reviver.call(t,e,o)}var j;if(text+=\"\",rx_dangerous.lastIndex=0,rx_dangerous.test(text)&&(text=text.replace(rx_dangerous,function(t){return\"\\\\u\"+(\"0000\"+t.charCodeAt(0).toString(16)).slice(-4)})),rx_one.test(text.replace(rx_two,\"@\").replace(rx_three,\"]\").replace(rx_four,\"\")))return j=eval(\"(\"+text+\")\"),\"function\"==typeof reviver?walk({\"\":j},\"\"):j;throw new SyntaxError(\"JSON.parse\")})}();"), CComVariant()));
	hr = hr;


	CComVariant varRes;
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsTest"), CComVariant(L"{a:1}"), CComVariant(), CComVariant(), &varRes));
	OutputDebugString(L"Test: ");
	OutputDebugVariant(varRes);
	OutputDebugString(L"\r\n");
	varRes.Clear();
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsTest"), CComVariant(123), CComVariant(), CComVariant(), &varRes));
	OutputDebugString(L"Test: ");
	OutputDebugVariant(varRes);
	OutputDebugString(L"\r\n");
	varRes.Clear();

	CComVariant varObj;
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsParse"), CComVariant(L"{\"a\":1}"), CComVariant(), CComVariant(), &varObj));
	OutputDebugString(L"Object: ");
	OutputDebugVariant(varObj);
	OutputDebugString(L"\r\n");

	CComVariant varTemp;
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsSetObjKey"), varObj, CComVariant(L"b"), CComVariant(3.1415), &varTemp));
	OutputDebugString(L"Temp: ");
	OutputDebugVariant(varTemp);
	OutputDebugString(L"\r\n");
	varTemp.Clear();
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsGetObjKey"), varObj, CComVariant(L"b"), CComVariant(), &varTemp));
	OutputDebugString(L"Temp: ");
	OutputDebugVariant(varTemp);
	OutputDebugString(L"\r\n");
	varTemp.Clear();
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsGetObjKey"), varObj, CComVariant(L"c"), CComVariant(), &varTemp));
	OutputDebugString(L"Temp: ");
	OutputDebugVariant(varTemp);
	OutputDebugString(L"\r\n");
	varTemp.Clear();

	CComVariant varStr;
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsStringify"), varObj, CComVariant(), CComVariant(2), &varStr));
	OutputDebugString(L"String: ");
	OutputDebugVariant(varStr);
	OutputDebugString(L"\r\n");

	CComBSTR bstrText;
	CHECKHR(EncodeURIComponent(CComBSTR(L"http://www.arcgis.com?q=abc"), &bstrText));
	OutputDebugString(L"Encode: ");
	OutputDebugString((BSTR) bstrText);
	OutputDebugString(L"\r\n");

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSJScript::EncodeURIComponent(BSTR bstrURI, BSTR* pbstrEncodedURI)
{
	HRESULT hr = S_OK;
	if (!pbstrEncodedURI) return E_INVALIDARG;
	*pbstrEncodedURI = NULL;
	if (!m_spIQSScriptSite) return E_POINTER;
	CComVariant varResult;
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsEncodeURIComponent"), CComVariant(bstrURI), CComVariant(), CComVariant(), &varResult));
	CHECKHR(VariantToBSTR(varResult, pbstrEncodedURI));
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
