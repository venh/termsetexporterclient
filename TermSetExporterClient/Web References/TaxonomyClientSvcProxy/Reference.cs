﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.1008
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

// 
// This source code was auto-generated by Microsoft.VSDesigner, Version 4.0.30319.1008.
// 
#pragma warning disable 1591

namespace TermSetExporterClient.TaxonomyClientSvcProxy {
    using System;
    using System.Web.Services;
    using System.Diagnostics;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Xml.Serialization;
    
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="Taxonomy web serviceSoap", Namespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/")]
    public partial class Taxonomywebservice : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        private System.Threading.SendOrPostCallback GetTermSetsOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetTermsByLabelOperationCompleted;
        
        private System.Threading.SendOrPostCallback AddTermsOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetKeywordTermsByGuidsOperationCompleted;
        
        private bool useDefaultCredentialsSetExplicitly;
        
        /// <remarks/>
        public Taxonomywebservice() {
            this.Url = global::TermSetExporterClient.Properties.Settings.Default.TermSetExporterClient_TaxonomyClientSvcProxy_Taxonomy_x0020_web_x0020_service;
            if ((this.IsLocalFileSystemWebService(this.Url) == true)) {
                this.UseDefaultCredentials = true;
                this.useDefaultCredentialsSetExplicitly = false;
            }
            else {
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        public new string Url {
            get {
                return base.Url;
            }
            set {
                if ((((this.IsLocalFileSystemWebService(base.Url) == true) 
                            && (this.useDefaultCredentialsSetExplicitly == false)) 
                            && (this.IsLocalFileSystemWebService(value) == false))) {
                    base.UseDefaultCredentials = false;
                }
                base.Url = value;
            }
        }
        
        public new bool UseDefaultCredentials {
            get {
                return base.UseDefaultCredentials;
            }
            set {
                base.UseDefaultCredentials = value;
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        /// <remarks/>
        public event GetTermSetsCompletedEventHandler GetTermSetsCompleted;
        
        /// <remarks/>
        public event GetTermsByLabelCompletedEventHandler GetTermsByLabelCompleted;
        
        /// <remarks/>
        public event AddTermsCompletedEventHandler AddTermsCompleted;
        
        /// <remarks/>
        public event GetKeywordTermsByGuidsCompletedEventHandler GetKeywordTermsByGuidsCompleted;
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://schemas.microsoft.com/sharepoint/taxonomy/soap/GetTermSets", RequestNamespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/", ResponseNamespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string GetTermSets(string sharedServiceIds, string termSetIds, int lcid, string clientTimeStamps, string clientVersions, out string serverTermSetTimeStampXml) {
            object[] results = this.Invoke("GetTermSets", new object[] {
                        sharedServiceIds,
                        termSetIds,
                        lcid,
                        clientTimeStamps,
                        clientVersions});
            serverTermSetTimeStampXml = ((string)(results[1]));
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void GetTermSetsAsync(string sharedServiceIds, string termSetIds, int lcid, string clientTimeStamps, string clientVersions) {
            this.GetTermSetsAsync(sharedServiceIds, termSetIds, lcid, clientTimeStamps, clientVersions, null);
        }
        
        /// <remarks/>
        public void GetTermSetsAsync(string sharedServiceIds, string termSetIds, int lcid, string clientTimeStamps, string clientVersions, object userState) {
            if ((this.GetTermSetsOperationCompleted == null)) {
                this.GetTermSetsOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetTermSetsOperationCompleted);
            }
            this.InvokeAsync("GetTermSets", new object[] {
                        sharedServiceIds,
                        termSetIds,
                        lcid,
                        clientTimeStamps,
                        clientVersions}, this.GetTermSetsOperationCompleted, userState);
        }
        
        private void OnGetTermSetsOperationCompleted(object arg) {
            if ((this.GetTermSetsCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetTermSetsCompleted(this, new GetTermSetsCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://schemas.microsoft.com/sharepoint/taxonomy/soap/GetTermsByLabel", RequestNamespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/", ResponseNamespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string GetTermsByLabel(string label, int lcid, StringMatchOption matchOption, int resultCollectionSize, string termIds, bool addIfNotFound) {
            object[] results = this.Invoke("GetTermsByLabel", new object[] {
                        label,
                        lcid,
                        matchOption,
                        resultCollectionSize,
                        termIds,
                        addIfNotFound});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void GetTermsByLabelAsync(string label, int lcid, StringMatchOption matchOption, int resultCollectionSize, string termIds, bool addIfNotFound) {
            this.GetTermsByLabelAsync(label, lcid, matchOption, resultCollectionSize, termIds, addIfNotFound, null);
        }
        
        /// <remarks/>
        public void GetTermsByLabelAsync(string label, int lcid, StringMatchOption matchOption, int resultCollectionSize, string termIds, bool addIfNotFound, object userState) {
            if ((this.GetTermsByLabelOperationCompleted == null)) {
                this.GetTermsByLabelOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetTermsByLabelOperationCompleted);
            }
            this.InvokeAsync("GetTermsByLabel", new object[] {
                        label,
                        lcid,
                        matchOption,
                        resultCollectionSize,
                        termIds,
                        addIfNotFound}, this.GetTermsByLabelOperationCompleted, userState);
        }
        
        private void OnGetTermsByLabelOperationCompleted(object arg) {
            if ((this.GetTermsByLabelCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetTermsByLabelCompleted(this, new GetTermsByLabelCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://schemas.microsoft.com/sharepoint/taxonomy/soap/AddTerms", RequestNamespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/", ResponseNamespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string AddTerms(System.Guid sharedServiceId, System.Guid termSetId, int lcid, string newTerms) {
            object[] results = this.Invoke("AddTerms", new object[] {
                        sharedServiceId,
                        termSetId,
                        lcid,
                        newTerms});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void AddTermsAsync(System.Guid sharedServiceId, System.Guid termSetId, int lcid, string newTerms) {
            this.AddTermsAsync(sharedServiceId, termSetId, lcid, newTerms, null);
        }
        
        /// <remarks/>
        public void AddTermsAsync(System.Guid sharedServiceId, System.Guid termSetId, int lcid, string newTerms, object userState) {
            if ((this.AddTermsOperationCompleted == null)) {
                this.AddTermsOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAddTermsOperationCompleted);
            }
            this.InvokeAsync("AddTerms", new object[] {
                        sharedServiceId,
                        termSetId,
                        lcid,
                        newTerms}, this.AddTermsOperationCompleted, userState);
        }
        
        private void OnAddTermsOperationCompleted(object arg) {
            if ((this.AddTermsCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.AddTermsCompleted(this, new AddTermsCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://schemas.microsoft.com/sharepoint/taxonomy/soap/GetKeywordTermsByGuids", RequestNamespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/", ResponseNamespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string GetKeywordTermsByGuids(string termIds, int lcid) {
            object[] results = this.Invoke("GetKeywordTermsByGuids", new object[] {
                        termIds,
                        lcid});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void GetKeywordTermsByGuidsAsync(string termIds, int lcid) {
            this.GetKeywordTermsByGuidsAsync(termIds, lcid, null);
        }
        
        /// <remarks/>
        public void GetKeywordTermsByGuidsAsync(string termIds, int lcid, object userState) {
            if ((this.GetKeywordTermsByGuidsOperationCompleted == null)) {
                this.GetKeywordTermsByGuidsOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetKeywordTermsByGuidsOperationCompleted);
            }
            this.InvokeAsync("GetKeywordTermsByGuids", new object[] {
                        termIds,
                        lcid}, this.GetKeywordTermsByGuidsOperationCompleted, userState);
        }
        
        private void OnGetKeywordTermsByGuidsOperationCompleted(object arg) {
            if ((this.GetKeywordTermsByGuidsCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetKeywordTermsByGuidsCompleted(this, new GetKeywordTermsByGuidsCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        public new void CancelAsync(object userState) {
            base.CancelAsync(userState);
        }
        
        private bool IsLocalFileSystemWebService(string url) {
            if (((url == null) 
                        || (url == string.Empty))) {
                return false;
            }
            System.Uri wsUri = new System.Uri(url);
            if (((wsUri.Port >= 1024) 
                        && (string.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) == 0))) {
                return true;
            }
            return false;
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.0.30319.1015")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.microsoft.com/sharepoint/taxonomy/soap/")]
    public enum StringMatchOption {
        
        /// <remarks/>
        StartsWith,
        
        /// <remarks/>
        ExactMatch,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    public delegate void GetTermSetsCompletedEventHandler(object sender, GetTermSetsCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetTermSetsCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetTermSetsCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
        
        /// <remarks/>
        public string serverTermSetTimeStampXml {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[1]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    public delegate void GetTermsByLabelCompletedEventHandler(object sender, GetTermsByLabelCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetTermsByLabelCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetTermsByLabelCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    public delegate void AddTermsCompletedEventHandler(object sender, AddTermsCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class AddTermsCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal AddTermsCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    public delegate void GetKeywordTermsByGuidsCompletedEventHandler(object sender, GetKeywordTermsByGuidsCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.1")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetKeywordTermsByGuidsCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetKeywordTermsByGuidsCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
}

#pragma warning restore 1591