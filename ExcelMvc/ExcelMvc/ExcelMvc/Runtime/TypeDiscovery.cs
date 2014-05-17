/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Developer:         Wolfgang Stamm, Germany

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or 
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING 
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, 
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This program is free software; you can redistribute it and/or modify it under the terms of the 
GNU General Public License as published by the Free Software Foundation; either version 2 of 
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program;
if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, 
Boston, MA 02110-1301 USA.
*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ExcelMvc.Extensions;

namespace ExcelMvc.Runtime
{
    internal class TypeDiscoveryProxy : MarshalByRefObject
    {
        public override Object InitializeLifetimeService()
        {
            /*
            ILease lease = (ILease)base.InitializeLifetimeService();
            if (lease.CurrentState == LeaseState.Initial)
            {
              //lease.InitialLeaseTime = TimeSpan.FromMinutes(1);
              //lease.SponsorshipTimeout = TimeSpan.FromMinutes(2);
              //lease.RenewOnCallTime = TimeSpan.FromSeconds(2);                             
            }
            return lease;
            */
            return null;    //null for long-live proxy
        }

        [Serializable]
        public class Result
        {
            public List<string> Types { get; set; }
            public Exception Error { get; set; }
        }

        public Result Discover(string assemblyPath, Type type)
        {
            var result = new Result {Types = new List<string>()};
            try
            {
                var asm = Assembly.LoadFrom(assemblyPath);
                foreach (var item in asm.GetTypes())
                {
                    if (type.IsAssignableFrom(item) && !item.IsInterface && !item.IsAbstract)
                        result.Types.Add(item.AssemblyQualifiedName);
                }
            }
            catch (Exception e)
            {
                result.Error = e;
            }
            return result;
        }
    }

    internal class TypeDiscoveryDomains
    {
        public static AppDomain CreateDomain(string domainName, string privateBinPath)
        {
            var domain = AppDomain.CurrentDomain;
            if (String.Compare(domainName, domain.FriendlyName, StringComparison.OrdinalIgnoreCase) != 0)
            {
                var ads = new AppDomainSetup
                {
                    ApplicationBase = privateBinPath ?? Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    DisallowBindingRedirects = false,
                    DisallowCodeDownload = true,
                    ConfigurationFile = domain.SetupInformation.ConfigurationFile,
                    PrivateBinPath = privateBinPath
                };
                domain = AppDomain.CreateDomain(domainName, null, ads);
            }
            return domain;
        }

        public static TypeDiscoveryProxy CreateProxy(AppDomain domain)
        {
            return domain.FriendlyName.CompareOrdinalIgnoreCase(AppDomain.CurrentDomain.FriendlyName) == 0
                ? new TypeDiscoveryProxy()
                : (TypeDiscoveryProxy) domain.CreateInstanceAndUnwrap(Assembly.GetExecutingAssembly().FullName, typeof (TypeDiscoveryProxy).FullName);
        }

        public static List<string> Discover(string assemblyPath, Type type)
        {
            var domain = CreateDomain("~~DiscoveryDomain~~", null);
            try
            {
                var proxy = CreateProxy(domain);
                var result = proxy.Discover(assemblyPath, type);
                if (result.Error != null)
                    throw result.Error;
                return result.Types;
            }
            finally
            {
                AppDomain.Unload(domain);
            }
        }
    }
}
