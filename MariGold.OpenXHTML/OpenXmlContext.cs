﻿namespace MariGold.OpenXHTML
{
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Collections.Generic;

    internal sealed class OpenXmlContext : IOpenXmlContext
    {
        private WordprocessingDocument document;
        private List<DocxElement> elements;
        private List<ITextElement> textElements;
        private Dictionary<Int16, AbstractNum> abstractNumList;
        private Dictionary<Int16, NumberingInstance> numberingInstanceList;
        private string imagePath;
        private string baseUrl;
        private string uriSchema;
        private IParser parser;
        private Int16 listNumberId = 1;
        public Int32 RelationshipId { get; set; }

        private void PrepareWordElements()
        {
            elements = new List<DocxElement>() {
                new DocxDiv(this),
                new DocxUL(this),
                new DocxOL(this),
                new DocxImage(this),
                new DocxSpan(this),
                new DocxA(this),
                new DocxBr(this),
                new DocxUnderline(this),
                new DocxCenter(this),
                new DocxItalic(this),
                new DocxBold(this),
                new DocxHeading(this),
                new DocxHeader(this),
                new DocxFooter(this),
                new DocxAddress(this),
                new DocxSection(this),
                new DocxFont(this),
                new DocxDL(this),
                new DocxHr(this),
                new DocxQ(this),
                new DocxSup(this),
                new DocxSub(this),
                new DocxStrike(this),
                new DocxObject(this),
                new DocxTable(this),
                new DocxInline(this)
            };

            textElements = new List<ITextElement>() {
                new DocxBold(this),
                new DocxSpan(this),
                new DocxBr(this),
                new DocxCenter(this),
                new DocxItalic(this),
                new DocxUnderline(this),
                new DocxImage(this),
                new DocxDiv(this),
                new DocxHeader(this),
                new DocxFooter(this),
                new DocxFont(this),
                new DocxQ(this),
                new DocxSup(this),
                new DocxSub(this),
                new DocxStrike(this),
                new DocxObject(this),
                new DocxInline(this)
            };
        }

        private void SaveNumberDefinitions()
        {
            if (abstractNumList != null && numberingInstanceList != null)
            {
                if (MainDocumentPart.NumberingDefinitionsPart == null)
                {
                    _ = MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("numberingDefinitionsPart");
                }

                Numbering numbering = new Numbering();

                foreach (var abstractNum in abstractNumList)
                {
                    numbering.Append(abstractNum.Value);
                }

                foreach (var numberingInstance in numberingInstanceList)
                {
                    numbering.Append(numberingInstance.Value);
                }

                MainDocumentPart.NumberingDefinitionsPart.Numbering = numbering;
            }
        }

        internal OpenXmlContext(WordprocessingDocument document)
        {
            this.document = document;
            if (this.document.MainDocumentPart == null)
            {
                var mainPart = this.document.AddMainDocumentPart();
                mainPart.Document = new Document();
            }

            PrepareWordElements();
        }

        public string ImagePath
        {
            get
            {
                return imagePath;
            }

            set
            {
                imagePath = value;
            }
        }

        public string BaseURL
        {
            get
            {
                return baseUrl;
            }

            set
            {
                baseUrl = value;
            }
        }

        public string UriSchema
        {
            get
            {
                return uriSchema;
            }

            set
            {
                uriSchema = value;
            }
        }

        public IParser Parser
        {
            get
            {
                return parser;
            }
        }

        public WordprocessingDocument WordprocessingDocument
        {
            get
            {
                if (document == null)
                {
                    throw new InvalidOperationException("Document is not opened!");
                }

                return document;
            }
        }

        public MainDocumentPart MainDocumentPart
        {
            get
            {
                var mainPart = document?.MainDocumentPart;
                if (mainPart == null)
                {
                    throw new InvalidOperationException("Document is not opened!");
                }

                return mainPart;
            }
        }

        public Document Document => MainDocumentPart.Document;

        public Int16 ListNumberId
        {
            get
            {
                return listNumberId;
            }

            set
            {
                listNumberId = value;
            }
        }

        public void Save()
        {
            SaveNumberDefinitions();

            Document.Save();
        }

        public DocxElement Convert(DocxNode node)
        {
            foreach (DocxElement element in elements)
            {
                if (element.CanConvert(node))
                {
                    return element;
                }
            }

            return null;
        }

        public ITextElement ConvertTextElement(DocxNode node)
        {
            foreach (ITextElement element in textElements)
            {
                if (element.CanConvert(node))
                {
                    return element;
                }
            }

            return null;
        }

        public DocxElement GetBodyElement()
        {
            return new DocxBody(this);
        }

        public void SaveNumberingDefinition(Int16 numberId, AbstractNum abstractNum, NumberingInstance numberingInstance)
        {
            if (abstractNumList == null)
            {
                abstractNumList = new Dictionary<Int16, AbstractNum>();
            }

            if (numberingInstanceList == null)
            {
                numberingInstanceList = new Dictionary<Int16, NumberingInstance>();
            }

            if (!abstractNumList.ContainsKey(numberId))
            {
                abstractNumList.Add(numberId, abstractNum);
            }

            if (!numberingInstanceList.ContainsKey(numberId))
            {
                numberingInstanceList.Add(numberId, numberingInstance);
            }
        }

        public void SetParser(IParser parser)
        {
            this.parser = parser ?? throw new ArgumentNullException("parser");
        }

        public IDocxInterchanger GetInterchanger()
        {
            return new DocxInterchanger();
        }

        public void Dispose()
        {
            document?.Close();
            document?.Dispose();

            document = null;
        }
    }
}
