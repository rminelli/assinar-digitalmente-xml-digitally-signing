using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Threading;

// -----------------------------------------------------------------------------------------------------------------
//                                         Biblioteca NF-eletrônica.
// -----------------------------------------------------------------------------------------------------------------
// 
// Condições de uso:
// 
// Este material pode ser copiado, distribuído, exibido e executado, bem como utilizado para criar obras derivadas,
// sob licença Creative Commons Attribution (http://creativecommons.org/licenses/by/2.5/br/), com a citação da fonte
// e o devido crédito.
//
//    
//                                      (c) 2019 - NF-eletronica.com
// -----------------------------------------------------------------------------------------------------------------
//
//   Nota: Para que o código funcione é necessário adicionar a referência “System.Security.Criptography.Xml” ao //         projeto do Visual Studio, da seguinte forma:
//
//     1. selecione “References” no Solution Explorer eclique com o botão direito; 
//
//     2. selecione “Add References…”; 
//
//     3. procure o componente “System.Security” na aba .NET e confirme; 
// -----------------------------------------------------------------------------------------------------------------
//
//      Atualização do tratamento de expressões relugares mantendo a formatação de e-mail.
//      
//      Roberto Minelli - 2019
//      


namespace AssinaXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string msgLog = "";

            string _arquivo = args[0];
            if ( _arquivo == null )
            {
                msgLog += "\rNome de arquivo não informado...";
            }
            else if (!File.Exists(_arquivo))
            {
                msgLog += "\rArquivo {0} inexistente..." + _arquivo;
            }
            else
            {
                //Console.Write("URI a ser assinada (Ex.: infCanc, infNFe, infInut, etc.) :");
                //string _uri = Console.ReadLine();
                string _uri = args[1]; //"LoteRps";
                if (_uri == null)
                {
                    msgLog += "\rURI não informada...";
                }
                else
                {
                    //
                    //   le o arquivo xml
                    //
                    StreamReader SR;
                    string _stringXml;
                    SR = File.OpenText(_arquivo);
                    _stringXml = RemoverCaracteresEspeciais(RemoverAcentos(SR.ReadToEnd()));
                    //_stringXml = SR.ReadToEnd();
                    SR.Close();
                    //
                    //  realiza assinatura
                    //
                    AssinaturaDigital AD = new AssinaturaDigital();
                    //
                    //  cria cert
                    //
                    X509Certificate2 cert = new X509Certificate2();
                    //
                    //  seleciona certificado do repositório MY do windows
                    //
                    //Certificado certificado = new Certificado();
                    //string nroSerieCertificado = Convert.ToString(ConfigurationManager.AppSettings["nroSerieCertificado"]);
                    //cert = certificado.BuscaNroSerie(nroSerieCertificado);
                    int resultado = AD.Assinar(_stringXml, _uri, args[2], args[3]);
                    if (resultado == 0)
                    {
                        //
                        //  grava arquivo assinado
                        //
                        
                        StreamWriter SW;
                        SW = File.CreateText(_arquivo.Replace(".xml","_ass.xml").Trim());
                        SW.Write(AD.XMLStringAssinado);
                        SW.Close();
                        
                        /*
                        XmlTextWriter writer = new XmlTextWriter(_arquivo.Replace(".xml", "_ass.xml").Trim(), null);
                        writer.Formatting = Formatting.Indented;
                        AD.XMLDocAssinado.Save(writer);
                        */
                    }

                    msgLog += AD.mensagemResultado;

                    //
                    //  grava arquivo de log
                    //
                    StreamWriter Log;
                    Log = File.CreateText(_arquivo.Replace(".xml", "_log.txt").Trim());
                    Log.Write(msgLog);
                    Log.Close();
                }
            }

        }

        /// <summary>
        /// Remove caracteres especiais de uma string.
        /// </summary>
        /// <param name="valor">String com caracteres especiais</param>
        /// <returns>Parâmetro 'valor' sem caracteres especiais</returns>
        public static string RemoverCaracteresEspeciais(string valor)
        {
            //Regex r = new Regex("[^-<>/_.,=\"#:| A-z0-9\t\n]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
            Regex r = new Regex(@"^[A-Za-z0-9](([_\.\-]?[a-zA-Z0-9]+)*)@([A-Za-z0-9]+)(([\.\-]?[a-zA-Z0-9]+)*)\.([A-Za-z]{2,})$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
            
            return r.Replace(valor, String.Empty);
        }

        /// <summary>
        /// Função para remover acentos da string
        /// </summary>
        /// <param name="valor">String para receber tratamento</param>
        /// <returns>String tratada</returns>
        public static string RemoverAcentos(string valor)
        {
            valor = Regex.Replace(valor, "[ÁÀÂÃ]", "A");
            valor = Regex.Replace(valor, "[ÉÈÊ]", "E");
            valor = Regex.Replace(valor, "[Í]", "I");
            valor = Regex.Replace(valor, "[ÓÒÔÕ]", "O");
            valor = Regex.Replace(valor, "[ÚÙÛÜ]", "U");
            valor = Regex.Replace(valor, "[Ç]", "C");
            valor = Regex.Replace(valor, "[áàâã]", "a");
            valor = Regex.Replace(valor, "[éèê]", "e");
            valor = Regex.Replace(valor, "[í]", "i");
            valor = Regex.Replace(valor, "[óòôõ]", "o");
            valor = Regex.Replace(valor, "[úùûü]", "u");
            valor = Regex.Replace(valor, "[ç]", "c");
            valor = Regex.Replace(valor, "º", "o");
            valor = Regex.Replace(valor, "ª", "a");
            return valor;
        }
    }
    public class AssinaturaDigital
    {
        public int Assinar(string XMLString, string RefUri, string nroSerieCertificado, string AssinaRPS)
        /*
         *     Entradas:
         *         XMLString: string XML a ser assinada
         *         RefUri   : Referência da URI a ser assinada (Ex. infNFe
         *         X509Cert : certificado digital a ser utilizado na assinatura digital
         * 
         *     Retornos:
         *         Assinar : 0 - Assinatura realizada com sucesso
         *                   1 - Erro: Problema ao acessar o certificado digital - %exceção%
         *                   2 - Problemas no certificado digital
         *                   3 - XML mal formado + exceção
         *                   4 - A tag de assinatura %RefUri% inexiste
         *                   5 - A tag de assinatura %RefUri% não é unica
         *                   6 - Erro Ao assinar o documento - ID deve ser string %RefUri(Atributo)%
         *                   7 - Erro: Ao assinar o documento - %exceção%
         * 
         *         XMLStringAssinado : string XML assinada
         * 
         *         XMLDocAssinado    : XMLDocument do XML assinado
         */
        {
            int resultado = 0;
            msgResultado = "Assinatura realizada com sucesso";
            try
            {
                //   certificado para ser utilizado na assinatura
                //
                string _xnome = "";

                //
                //  cria X509Cert
                //
                X509Certificate2 X509Cert = new X509Certificate2();
                //
                //  seleciona certificado do repositório MY do windows
                //
                Certificado certificado = new Certificado();
                X509Cert = certificado.BuscaNroSerie(nroSerieCertificado);

                if (X509Cert != null)
                {
                    _xnome = X509Cert.Subject.ToString();
                }
                X509Certificate2 _X509Cert = new X509Certificate2();
                X509Store store = new X509Store("MY", StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
                X509Certificate2Collection collection = (X509Certificate2Collection)store.Certificates;
                X509Certificate2Collection collection1 = (X509Certificate2Collection)collection.Find(X509FindType.FindBySubjectDistinguishedName, _xnome, false);
                if (collection1.Count == 0)
                {
                    resultado = 2;
                    msgResultado = "Problemas no certificado digital";
                }
                else
                {
                    // certificado ok
                    _X509Cert = collection1[0];
                    string x;
                    x = _X509Cert.GetKeyAlgorithm().ToString();
                    // Create a new XML document.
                    XmlDocument doc = new XmlDocument();

                    // Format the document to ignore white spaces.
                    doc.PreserveWhitespace = false;

                    // Load the passed XML file using it's name.
                    try
                    {
                        doc.LoadXml(XMLString);

                        // Verifica se a tag a ser assinada existe é única
                        int qtdeRefUri = doc.GetElementsByTagName(RefUri).Count;

                        if (qtdeRefUri == 0)
                        {
                            //  a URI indicada não existe
                            resultado = 4;
                            msgResultado = "A tag de assinatura " + RefUri.Trim() + " inexiste";
                        }
                        // Exsiste mais de uma tag a ser assinada
                        else
                        {
                            if (qtdeRefUri > 1)
                            {
                                // existe mais de uma URI indicada
                                resultado = 5;
                                msgResultado = "A tag de assinatura " + RefUri.Trim() + " não é unica";

                            }
                            //else if (_listaNum.IndexOf(doc.GetElementsByTagName(RefUri).Item(0).Attributes.ToString().Substring(1,1))>0)
                            //{
                            //    resultado = 6;
                            //    msgResultado = "Erro: Ao assinar o documento - ID deve ser string (" + doc.GetElementsByTagName(RefUri).Item(0).Attributes + ")";
                            //}
                            else
                            {
                                try
                                {

                                    // Create a SignedXml object.
                                    SignedXml signedXml = new SignedXml(doc);

                                    // Add the key to the SignedXml document 
                                    signedXml.SigningKey = _X509Cert.PrivateKey;

                                    // Create a reference to be signed
                                    Reference reference = new Reference();

                                    XmlAttributeCollection _Uri;
                                    XmlDsigEnvelopedSignatureTransform env;
                                    XmlDsigC14NTransform c14;
                                    KeyInfo keyInfo;
                                    XmlElement xmlSignature, xmlSignedInfo, xmlKeyInfo, xmlSignatureValue;
                                    XmlAttribute attr;
                                    string signBase64;
                                    XmlText text;

                                    if (AssinaRPS == "true")
                                    {
                                        // pega o uri que deve ser assinada
                                        _Uri = doc.GetElementsByTagName("InfDeclaracaoPrestacaoServico").Item(0).Attributes;
                                        foreach (XmlAttribute _atributo in _Uri)
                                            if (_atributo.Name == "Id")
                                                reference.Uri = "#" + _atributo.InnerText;

                                        // Add an enveloped transformation to the reference.
                                        env = new XmlDsigEnvelopedSignatureTransform();
                                        reference.AddTransform(env);

                                        c14 = new XmlDsigC14NTransform();
                                        reference.AddTransform(c14);

                                        // Add the reference to the SignedXml object.
                                        signedXml.AddReference(reference);

                                        // Create a new KeyInfo object
                                        keyInfo = new KeyInfo();

                                        // Load the certificate into a KeyInfoX509Data object
                                        // and add it to the KeyInfo object.
                                        keyInfo.AddClause(new KeyInfoX509Data(_X509Cert));

                                        // Add the KeyInfo object to the SignedXml object.
                                        signedXml.KeyInfo = keyInfo;

                                        signedXml.ComputeSignature();

                                        xmlSignature = doc.CreateElement("Signature", "http://www.w3.org/2000/09/xmldsig#");

                                        attr = doc.CreateAttribute("Id");
                                        attr.Value = "Ass_" + reference.Uri.Replace("#", "");

                                        xmlSignature.Attributes.SetNamedItem(attr);

                                        xmlSignedInfo = signedXml.SignedInfo.GetXml();
                                        xmlKeyInfo = signedXml.KeyInfo.GetXml();

                                        xmlSignatureValue = doc.CreateElement("SignatureValue", xmlSignature.NamespaceURI);
                                        signBase64 = Convert.ToBase64String(signedXml.Signature.SignatureValue);
                                        text = doc.CreateTextNode(signBase64);
                                        xmlSignatureValue.AppendChild(text);

                                        xmlSignature.AppendChild(doc.ImportNode(xmlSignedInfo, true));
                                        xmlSignature.AppendChild(xmlSignatureValue);
                                        xmlSignature.AppendChild(doc.ImportNode(xmlKeyInfo, true));

                                        doc.GetElementsByTagName("Rps").Item(0).AppendChild(xmlSignature);
                                    }

                                    _Uri = doc.GetElementsByTagName("LoteRps").Item(0).Attributes;
                                    foreach (XmlAttribute _atributo in _Uri)
                                        if (_atributo.Name == "Id")
                                            reference.Uri = "#" + _atributo.InnerText;

                                    signedXml.ComputeSignature();

                                    // Get the XML representation of the signature and save
                                    // it to an XmlElement object.
                                    xmlSignature = signedXml.GetXml();

                                    // Append the element to the XML document.
                                    doc.DocumentElement.AppendChild(doc.ImportNode(xmlSignature, true));
                                    XMLDoc = new XmlDocument();
                                    XMLDoc.PreserveWhitespace = false;
                                    XMLDoc = doc;
                                }
                                catch (Exception caught)
                                {
                                    resultado = 7;
                                    msgResultado = "Erro: Ao assinar o documento - " + caught.Message;
                                }
                            }
                        }
                    }
                    catch (Exception caught)
                    {
                        resultado = 3;
                        msgResultado = "Erro: XML mal formado - " + caught.Message;
                    }
                }
            }
            catch (Exception caught)
            {
                resultado = 1;
                msgResultado = "Erro: Problema ao acessar o certificado digital" + caught.Message;
            }

            return resultado;
        }

        //
        // mensagem de Retorno
        //
        private string msgResultado;
        private XmlDocument XMLDoc;

        public XmlDocument XMLDocAssinado
        {
            get { return XMLDoc; }
        }

        public string XMLStringAssinado
        {
            get { return XMLDoc.OuterXml; }
        }

        public string mensagemResultado
        {
            get { return msgResultado; }
        }


    }
    public class Certificado
    {
        public X509Certificate2 BuscaNome(string Nome)
        {
            X509Certificate2 _X509Cert = new X509Certificate2();
            try
            {

                X509Store store = new X509Store("MY", StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
                X509Certificate2Collection collection = (X509Certificate2Collection)store.Certificates;
                X509Certificate2Collection collection1 = (X509Certificate2Collection)collection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);
                X509Certificate2Collection collection2 = (X509Certificate2Collection)collection.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, false);
                if (Nome == "")
                {
                    X509Certificate2Collection scollection = X509Certificate2UI.SelectFromCollection(collection2, "Certificado(s) Digital(is) disponível(is)", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection);
                    if (scollection.Count == 0)
                    {
                        _X509Cert.Reset();
                        Console.WriteLine("Nenhum certificado escolhido", "Atenção");
                    }
                    else
                    {
                        _X509Cert = scollection[0];
                    }
                }
                else
                {
                    X509Certificate2Collection scollection = (X509Certificate2Collection)collection2.Find(X509FindType.FindBySubjectDistinguishedName, Nome, false);
                    if (scollection.Count == 0)
                    {
                        Console.WriteLine("Nenhum certificado válido foi encontrado com o nome informado: " + Nome, "Atenção");
                        _X509Cert.Reset();
                    }
                    else
                    {
                        _X509Cert = scollection[0];
                    }
                }
                store.Close();
                return _X509Cert;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                return _X509Cert;
            }
        }
        public X509Certificate2 BuscaNroSerie(string NroSerie)
        {
            X509Certificate2 _X509Cert = new X509Certificate2();
            try
            {

                X509Store store = new X509Store("MY", StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
                X509Certificate2Collection collection = (X509Certificate2Collection)store.Certificates;
                X509Certificate2Collection collection1 = (X509Certificate2Collection)collection.Find(X509FindType.FindByTimeValid, DateTime.Now, true);
                X509Certificate2Collection collection2 = (X509Certificate2Collection)collection1.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, true);
                if (NroSerie == "")
                {
                    X509Certificate2Collection scollection = X509Certificate2UI.SelectFromCollection(collection2, "Certificados Digitais", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection);
                    if (scollection.Count == 0)
                    {
                        _X509Cert.Reset();
                        Console.WriteLine("Nenhum certificado válido foi encontrado com o número de série informado: " + NroSerie, "Atenção");
                    }
                    else
                    {
                        _X509Cert = scollection[0];
                    }
                }
                else
                {
                    X509Certificate2Collection scollection = (X509Certificate2Collection)collection2.Find(X509FindType.FindBySerialNumber, NroSerie, true);
                    if (scollection.Count == 0)
                    {
                        _X509Cert.Reset();
                        Console.WriteLine("Nenhum certificado válido foi encontrado com o número de série informado: " + NroSerie, "Atenção");
                    }
                    else
                    {
                        _X509Cert = scollection[0];
                    }
                }
                store.Close();
                return _X509Cert;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                return _X509Cert;
            }
        }

    }
}
