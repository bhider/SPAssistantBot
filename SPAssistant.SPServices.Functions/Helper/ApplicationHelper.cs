﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;

namespace SPAssistant.SPServices.Functions.Helper
{
    //Note: This is required to resolve some of the dependency issues - certain dlls depend on a specific version of a dependent dll being
    //available. This sets up assembly redirection so that the references are correctly resolved - currently required for Newtonsoft.Json and
    //Microsoft.IdentityModel.ActiveDirectory.Client

    public static class AssemblyBindingRedirectHelper
    {

        ///<summary>
        /// Reads the "BindingRedirecs" field from the app settings and applies the redirection on the
        /// specified assemblies
        /// </summary>

        public static void ConfigureBindingRedirects()
        {
            var redirects = GetBindingRedirects();
            redirects.ForEach(RedirectAssembly);
        }

        private static List<BindingRedirect> GetBindingRedirects()
        {
            var result = new List<BindingRedirect>();
            var bindingRedirectListJson = Environment.GetEnvironmentVariable("BindingRedirects");
            using (var memoryStream = new MemoryStream(Encoding.Unicode.GetBytes(bindingRedirectListJson)))
            {
                var serializer = new DataContractJsonSerializer(typeof(List<BindingRedirect>));
                result = (List<BindingRedirect>)serializer.ReadObject(memoryStream);
            }
            return result;
        }

        private static void RedirectAssembly(BindingRedirect bindingRedirect)
        {
            ResolveEventHandler handler = null;
            handler = (sender, args) =>
            {
                var requestedAssembly = new AssemblyName(args.Name);
                if (requestedAssembly.Name != bindingRedirect.ShortName)
                {
                    return null;
                }
                var targetPublicKeyToken = new AssemblyName("x, PublicKeyToken=" + bindingRedirect.PublicKeyToken).GetPublicKeyToken();
                requestedAssembly.SetPublicKeyToken(targetPublicKeyToken);
                requestedAssembly.Version = new Version(bindingRedirect.RedirectToVersion);
                requestedAssembly.CultureInfo = CultureInfo.InvariantCulture;
                AppDomain.CurrentDomain.AssemblyResolve -= handler;
                var assembly = Assembly.Load(requestedAssembly);
                return assembly;
            };
            AppDomain.CurrentDomain.AssemblyResolve += handler;
        }

        public class BindingRedirect
        {
            public string ShortName { get; set; }
            public string PublicKeyToken { get; set; }
            public string RedirectToVersion { get; set; }
        }
    }

    public static class ApplicationHelper
    {
        private static bool IsStarted = false;
        private static object _syncLock = new object();
        ///<summary>
        /// Sets up the app before running any other code
        /// </summary>

        public static void Startup()
        {
            if (!IsStarted)
            {
                lock (_syncLock)
                {
                    if (!IsStarted)
                    {
                        AssemblyBindingRedirectHelper.ConfigureBindingRedirects();
                        IsStarted = true;
                    }
                }
            }
        }
    }
}
