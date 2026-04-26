namespace ReportForge.AddIn
{
    /// <summary>Ribbon XML 定义——定义 Word 功能区按钮布局</summary>
    public static class RibbonXml
    {
        public static string GetRibbonXml()
        {
            return @"<?xml version=""1.0"" encoding=""utf-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""rfTab"" label=""报告格式化"" insertAfterMso=""TabView"">

        <!-- 规范管理 -->
        <group id=""rfGroupSpec"" label=""规范管理"">
          <button id=""rfInitStyles"" label=""初始化样式""
                  size=""large"" imageMso=""StylesPane""
                  onAction=""OnInitializeStyles""
                  screentip=""根据配置在文档中创建9级样式体系和多级编号"" />
        </group>

        <!-- 样式应用 -->
        <group id=""rfGroupStyle"" label=""样式应用"">
          <button id=""rfBodyText"" label=""正文""
                  size=""normal"" imageMso=""ParagraphDialog""
                  onAction=""OnApplyBodyStyle""
                  screentip=""将选区设置为正文样式（仿宋四号）"" />
          <separator id=""rfSep1"" />
          <button id=""rfH1"" label=""一级标题"" tag=""1""
                  size=""normal"" imageMso=""OutlinePromote""
                  onAction=""OnApplyHeadingStyle"" screentip=""黑体三号 一、"" />
          <button id=""rfH2"" label=""二级标题"" tag=""2""
                  size=""normal"" imageMso=""OutlineDemote""
                  onAction=""OnApplyHeadingStyle"" screentip=""楷体三号 （一）"" />
          <button id=""rfH3"" label=""三级标题"" tag=""3""
                  size=""normal"" imageMso=""OutlineDemote""
                  onAction=""OnApplyHeadingStyle"" screentip=""仿宋四号粗 1."" />
          <button id=""rfH4"" label=""四级标题"" tag=""4""
                  size=""normal"" imageMso=""OutlineDemote""
                  onAction=""OnApplyHeadingStyle"" screentip=""仿宋四号粗 1)"" />
          <button id=""rfH5"" label=""五级"" tag=""5""
                  size=""normal"" imageMso=""OutlineDemote""
                  onAction=""OnApplyHeadingStyle"" screentip=""仿宋四号 ①"" />
          <button id=""rfH6"" label=""六级"" tag=""6""
                  size=""normal"" imageMso=""OutlineDemote""
                  onAction=""OnApplyHeadingStyle"" screentip=""仿宋四号 a)"" />
          <separator id=""rfSep2"" />
          <button id=""rfFigCaption"" label=""图题""
                  size=""normal"" imageMso=""CaptionInsert""
                  onAction=""OnInsertFigureCaption""
                  screentip=""在选区下方插入图题（原生题注）"" />
          <button id=""rfTblCaption"" label=""表题""
                  size=""normal"" imageMso=""CaptionInsert""
                  onAction=""OnInsertTableCaption""
                  screentip=""在选区上方插入表题（原生题注）"" />
          <button id=""rfStdTable"" label=""标准表格""
                  size=""normal"" imageMso=""TableInsert""
                  onAction=""OnApplyStandardTable""
                  screentip=""对选中表格应用标准格式（外框1.5磅、黑体小四表头）"" />
          <separator id=""rfSep3"" />
          <button id=""rfFigurePara"" label=""图片格式""
                  size=""normal"" imageMso=""PictureInsertFromFile""
                  onAction=""OnApplyFigureStyle""
                  screentip=""将选区设为图片段落样式（居中、无缩进）"" />
        </group>

        <!-- 文档结构 -->
        <group id=""rfGroupStructure"" label=""文档结构"">
          <button id=""rfPageSetup"" label=""页面设置""
                  size=""large"" imageMso=""PageSetupDialog""
                  onAction=""OnPageSetup""
                  screentip=""设置纸张A4、页边距(上下3cm/左右2.6cm)、页眉1.5cm/页脚2cm"" />
          <button id=""rfInsertTOC"" label=""插入规范目录""
                  size=""large"" imageMso=""TableOfContentsInsertClassic""
                  onAction=""OnInsertTOC""
                  screentip=""在光标位置插入符合格式要求的自动目录"" />
          <button id=""rfUpdateTOC"" label=""更新目录""
                  size=""large"" imageMso=""TableOfContentsUpdateClassic""
                  onAction=""OnUpdateTOC"" />
          <button id=""rfSetupHF"" label=""设置页眉页脚""
                  size=""large"" imageMso=""HeaderFooterInsert""
                  onAction=""OnSetupHeaderFooter""
                  screentip=""按节配置页眉页脚格式（封面无/目录罗马页码/正文页眉+页码）"" />
        </group>

        <!-- 修复与巡检 -->
        <group id=""rfGroupRepair"" label=""修复与巡检"">
          <button id=""rfSmartFormat"" label=""智能格式化""
                  size=""large"" imageMso=""AutoFormat""
                  onAction=""OnSmartFormat""
                  screentip=""自动识别全文标题层级并批量应用样式（支持多种编号体系）"" />
          <button id=""rfFullRepair"" label=""全文修复""
                  size=""large"" imageMso=""ReviewTrackChanges""
                  onAction=""OnFullRepair""
                  screentip=""清理直接格式 → 重建编号 → 更新字段"" />
          <button id=""rfInspect"" label=""格式巡检""
                  size=""large"" imageMso=""ReviewCompareMenu""
                  onAction=""OnRunInspection""
                  screentip=""扫描手动编号、样式偏差、层级跳级等问题"" />
          <button id=""rfClearDF"" label=""清理直接格式""
                  size=""normal"" imageMso=""ClearFormatting""
                  onAction=""OnClearDirectFormatting""
                  screentip=""清除选区/全文的字体和段落直接格式（跳过目录）"" />
        </group>

        <!-- 配置与帮助 -->
        <group id=""rfGroupConfig"" label=""配置与帮助"">
          <button id=""rfOpenConfig"" label=""打开配置文件""
                  size=""large"" imageMso=""FileOpen""
                  onAction=""OnOpenConfig""
                  screentip=""用记事本打开格式配置文件（修改后点重载生效）"" />
          <button id=""rfReloadConfig"" label=""重载配置""
                  size=""normal"" imageMso=""Refresh""
                  onAction=""OnReloadConfig""
                  screentip=""重新读取配置文件（修改JSON后使用）"" />
          <button id=""rfAbout"" label=""关于""
                  size=""normal"" imageMso=""Info""
                  onAction=""OnAbout""
                  screentip=""版本信息与版权声明"" />
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }
    }
}
