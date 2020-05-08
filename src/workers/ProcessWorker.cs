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
            TableService tableService, MainFormOption mainFormOption)
        {
            string output = mainFormOption.OutPutFolder.Text;

            if (mainFormOption.ExtractImageCheckBox.Checked)
            {
                imageService.extractImages(document, output);
            }
            if (mainFormOption.ExtractHyperLinkCheckBox.Checked)
            {
                hyperLinkService.extractHyperLink(document, output);
            }
            if (mainFormOption.ExtractTable.Checked)
            {
                tableService.extractTable(document, output);
            }
        }

        /*private void ParagraphSet(DocX document)
        {
            this.paragraphService.Set(document, SpaceBefore, SpaceAfter, SpaceLineVal, IndentationSpecialVal, IndentationBefore, IndentationAfter, TextSpace
                , ParagraphAlign, SpaceLineType, IndentationSpecial);
        }*/

        public static void TextReplaceSet(DocX document, TextReplaceService textReplaceService, MainFormOption mainFormOption)
        {
            DataGridView ReplaceTextGridView = mainFormOption.ReplaceTextGridView;
            DataGridView ReplaceLinkGridView = mainFormOption.ReplaceLinkGridView;
            Dictionary<string, string> replaceTextLists = getTextList(ReplaceTextGridView);
            textReplaceService.TextReplaceSet(document, replaceTextLists);
            //Dictionary<string, string> hyperLinkLists = getTextList(ReplaceTextGridView);
            //this.textReplaceService.HyperLinkReplaceSet(document, hyperLinkLists);
        }
        public static void PageSet(DocX document, PageSettingService pageSettingService, MainFormOption mainFormOption)
        {
            pageSettingService.marginSetting(document, mainFormOption.NotSetMargin.Checked, mainFormOption.TopMargin.Value.ToString(), mainFormOption.BottomMargin.Value.ToString(), mainFormOption.LeftMargin.Value.ToString(), mainFormOption.RightMargin.Value.ToString());
            pageSettingService.pageSizeSetting(document, mainFormOption.NotSetPageSize.Checked, mainFormOption.PageWidth.Value.ToString(), mainFormOption.PageHeight.Value.ToString());
            pageSettingService.pageOrientation(document, mainFormOption.PageSetOrientation.Text);
        }

        public static void HeaderFooterSet(DocX document, HeaderFooterService headerFooterService, MainFormOption mainFormOption)
        {
            if (mainFormOption.ClearHeader.Checked)
            {
                headerFooterService.clearHeader(document);
            }
            if (mainFormOption.ClearFooter.Checked)
            {
                headerFooterService.clearFooter(document);
            }

            Boolean firstOption = mainFormOption.FirstHeaderFooter.Checked;
            Boolean oddEvenOption = mainFormOption.OddEvenHeaderFooter.Checked;
            if (!mainFormOption.NotSetHeader.Checked && !mainFormOption.ClearHeader.Checked)
            {
                Font headerFont = mainFormOption.HeaderFontDialog.Font;
                string headerAlign = mainFormOption.HeaderAlignComBox.Text;
                Color headerColor = mainFormOption.HeaderColorDialog.Color;
                string pageHeaderText = mainFormOption.PageHeader.Text;
                string firstHeaderText = mainFormOption.FirstHeader.Text;
                string oddHeaderText = mainFormOption.OddHeader.Text;
                string evenHeaderText = mainFormOption.EvenHeader.Text;
                string headerImage = mainFormOption.HeaderImagePath.Text;
                Boolean headerLineBool = mainFormOption.HeaderLine.Checked;

                HeaderFooterOption headerOption = new HeaderFooterOption(headerFont, headerColor, headerAlign, pageHeaderText, firstHeaderText, oddHeaderText, evenHeaderText, headerImage, "", headerLineBool);

                headerFooterService.addHeaders(document, headerOption, firstOption, oddEvenOption);
            }

            if (!mainFormOption.NotSetFooter.Checked && !mainFormOption.ClearFooter.Checked)
            {
                Font footerFont = mainFormOption.FooterFontDialog.Font;
                string footerAlign = mainFormOption.FooterAlignComBox.Text;
                Color footerColor = mainFormOption.FooterColorDialog.Color;
                string pageFooterText = mainFormOption.PageFooter.Text;
                string firstFooterText = mainFormOption.FirstFooter.Text;
                string oddFooterText = mainFormOption.OddFooter.Text;
                string evenFooterText = mainFormOption.EvenFooter.Text;
                string footerImage = mainFormOption.FooterImagePath.Text;
                string pageNumber = mainFormOption.PageNumberComBox.Text;
                Boolean footerLineBool = mainFormOption.FooterLine.Checked;
                HeaderFooterOption footerOption = new HeaderFooterOption(footerFont, footerColor, footerAlign, pageFooterText, firstFooterText, oddFooterText, evenFooterText, footerImage, pageNumber, footerLineBool);

                headerFooterService.addFooters(document, footerOption, firstOption, oddEvenOption);
            }
        }

        public static void DocInfoSet(DocX document, DocInfoService docInfoService, MainFormOption mainFormOption)
        {
            string title = mainFormOption.DocTitle.Text;
            string subject = mainFormOption.DocSubject.Text;
            string category = mainFormOption.DocCategory.Text;
            string description = mainFormOption.DocDescription.Text;
            string creator = mainFormOption.DocCreator.Text;
            string version = mainFormOption.DocVersion.Text;
            Boolean editProtect = mainFormOption.DocEditPrctCheckBox.Checked;
            Boolean removeEditPrct = mainFormOption.DocEditPrctRemove.Checked;
            string editPassword = mainFormOption.DocEditPassword.Text;

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

        public static void UpdateFileTime(DocInfoService docInfoService, string targetFile, MainFormOption mainFormOption)
        {
            docInfoService.UpdateFileTime(targetFile, mainFormOption.CreateTimeCheckBox.Checked, mainFormOption.DocCreateTime.Value, mainFormOption.UpdateTimeCheckBox.Checked, mainFormOption.DocUpdateTime.Value);
        }

        public static Dictionary<string, string> getTextList(DataGridView ReplaceTextGridView)
        {
            Dictionary<string, string> lists = new Dictionary<string, string>();

            foreach (DataGridViewRow row in ReplaceTextGridView.Rows)
            {
                string source = row.Cells[0].Value == null ? "" : row.Cells[0].Value.ToString();
                string target = row.Cells[1].Value == null ? "" : row.Cells[1].Value.ToString();
                if (source.Length == 0 && target.Length == 0)
                {
                    continue;
                }
                lists.Add(source, target);
            }
            return lists;
        }
    }
}
