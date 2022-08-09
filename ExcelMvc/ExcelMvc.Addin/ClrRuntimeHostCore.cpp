/****************************** Module Header ******************************\
* Module Name:  ClrRuntimeHost.Core.cpp
* Copyright (c) Microsoft Corporation.
*
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/en-us/openness/licenses.aspx#MPL.
* All other rights reserved.
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include "pch.h"


BOOL 
ClrRuntimeHostCore::Start(PCWSTR pszVersion, PCWSTR pszAssemblyName)
{
	return FALSE;
}

void
ClrRuntimeHostCore::CallStaticMethod(PCWSTR pszClassName, PCWSTR pszMethodName, VARIANT *pArg1, VARIANT *pArg2, VARIANT *pArg3)
{
}

void
ClrRuntimeHostCore::Stop()
{
}