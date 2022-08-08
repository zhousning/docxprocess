using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Docx.src.model;
using Docx.src.services;
using Xceed.Words.NET;

namespace Docx.src.workers
{
    class ProcessWorker
    {
        public static void ExtractSet(DocX document, ImageService imageService, HyperLinkService hyperLinkService,
            TableService tableService, FormValOption formValOption)
        {
            string output = formValOption.OutPutFolder;

            if (formValOption.ExtractImageCheckBox)
            {
                imageService.extractImages(document, output);
            }
            if (formValOption.ExtractHyperLinkCheckBox)
            {
                hyperLinkService.extractHyperLink(document, output);
            }
            if (formValOption.ExtractTable)
            {
                tableService.extractTable(document, output);
            }
        }

        /*private void ParagraphSet(DocX document)
        {
            this.paragraphService.Set(document, SpaceBefore, SpaceAfter, SpaceLineVal, IndentationSpecialVal, IndentationBefore, IndentationAfter, TextSpace
                , ParagraphAlign, SpaceLineType, IndentationSpecial);
        }*/

        public static void TextReplaceSet(DocX document, TextReplaceService textReplaceService, FormValOption formValOption)
        {
            Dictionary<string, string> replaceTextLists = formValOption.ReplaceTextGridView;
            textReplaceService.TextReplaceSet(document, replaceTextLists);
            //Dictionary<string, string> replaceLinkGridView = formValOption.ReplaceLinkGridView;
            //thisReplaceService.HyperLinkReplaceSet(document, hyperLinkLists);
        }
        public static void PageSet(DocX document, PageSettingService pageSettingService, FormValOption formValOption)
        {
            pageSettingService.marginSetting(document, formValOption.NotSetMargin, formValOption.TopMargin, formValOption.BottomMargin, formValOption.LeftMargin, formValOption.RightMargin);
            pageSettingService.pageSizeSetting(document, formValOption.NotSetPageSize, formValOption.PageWidth, formValOption.PageHeight);
            pageSettingService.pageOrientation(document, formValOption.PageSetOrientation);
        }

        public static void HeaderFooterSet(DocX document, HeaderFooterService headerFooterService, FormValOption formValOption)
        {
            if (formValOption.ClearHeader)
            {
                headerFooterService.clearHeader(document);
            }
            if (formValOption.ClearFooter)
            {
                headerFooterService.clearFooter(document);
            }

            Boolean firstOption = formValOption.FirstHeaderFooter;
            Boolean oddEvenOption = formValOption.OddEvenHeaderFooter;
            if (!formValOption.NotSetHeader && !formValOption.ClearHeader)
            {
                Font headerFont = formValOption.HeaderFontDialog;
                string headerAlign = formValOption.HeaderAlignComBox;
                Color headerColor = formValOption.HeaderColorDialog;
                string pageHeaderText = formValOption.PageHeader;
                string firstHeaderText = formValOption.FirstHeader;
                string oddHeaderText = formValOption.OddHeader;
                string evenHeaderText = formValOption.EvenHeader;
                string headerImage = formValOption.HeaderImagePath;
                Boolean headerLineBool = formValOption.HeaderLine;
                HeaderFooterOption headerOption = new HeaderFooterOption(headerFont, headerColor, headerAlign, pageHeaderText, firstHeaderText, oddHeaderText, evenHeaderText, headerImage, "", headerLineBool);
                headerFooterService.clearHeader(document);
                headerFooterService.addHeaders(document, headerOption, firstOption, oddEvenOption);
            }

            if (!formValOption.NotSetFooter && !formValOption.ClearFooter)
            {
                Font footerFont = formValOption.FooterFontDialog;
                string footerAlign = formValOption.FooterAlignComBox;
                Color footerColor = formValOption.FooterColorDialog;
                string pageFooterText = formValOption.PageFooter;
                string firstFooterText = formValOption.FirstFooter;
                string oddFooterText = formValOption.OddFooter;
                string evenFooterText = formValOption.EvenFooter;
                string footerImage = formValOption.FooterImagePath;
                string pageNumber = formValOption.PageNumberComBox;
                Boolean footerLineBool = formValOption.FooterLine;
                HeaderFooterOption footerOption = new HeaderFooterOption(footerFont, footerColor, footerAlign, pageFooterText, firstFooterText, oddFooterText, evenFooterText, footerImage, pageNumber, footerLineBool);
                headerFooterService.clearFooter(document);
                headerFooterService.addFooters(document, footerOption, firstOption, oddEvenOption);
            }
        }

        public static void DocInfoSet(DocX document, DocInfoService docInfoService, FormValOption formValOption)
        {
            string title = formValOption.DocTitle;
            string subject = formValOption.DocSubject;
            string category = formValOption.DocCategory;
            string description = formValOption.DocDescription;
            string creator = formValOption.DocCreator;
            string version = formValOption.DocVersion;
            Boolean editProtect = formValOption.DocEditPrctCheckBox;
            Boolean removeEditPrct = formValOption.DocEditPrctRemove;
            string editPassword = formValOption.DocEditPassword;

            DocInfoOption option = new DocInfoOption(subject, title, creator, "", description, "", "", category, version, "", "");
            docInfoService.addCoreProperties(document, option);
            if (removeEditPrct)
            {
                docInfoService.DocRemoveProtect(document, removeEditPrct);
            }
            else if (editProtect)
            {
                docInfoService.DocProtect(document, editProtect, editPassword);
            }
        }

        public static void UpdateFileTime(DocInfoService docInfoService, string targetFile, FormValOption formValOption)
        {
            docInfoService.UpdateFileTime(targetFile, formValOption.CreateTimeCheckBox, formValOption.DocCreateTime, formValOption.UpdateTimeCheckBox, formValOption.DocUpdateTime);
        }

       
    }
}
