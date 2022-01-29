<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:tan="tag:textalign.net,2015:ns" xmlns:xs="http://www.w3.org/2001/XMLSchema"
   xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
   xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" exclude-result-prefixes="#all"
   expand-text="yes" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="3.0">

   <!-- Primary input: any XML file -->
   <!-- Secondary input: a docx template document -->
   <!-- Primary output: perhaps diagnostics -->
   <!-- Secondary output: the XML file serialized in a docx, saved at the specified target location;
      output will be indented to reflect the tree hierarchy of the input XML. See accompanying 
      README.md for further documentation, and read the comments below. -->

   <!-- This application was written to assist jTEI editors to allow jTEI XML to go through the copyediting 
      process. You ask, what does jTEI mean? Glad you asked. We're the world's premiere academic journal for 
      the study of scholarly text processing. Check us out: https://journal.tei-c.org/index.php/journal -->

   <!-- Version: 2021-01-29 -->
   <!-- Author: Joel Kalvesmaki -->
   <!-- License: GNU General Public License https://opensource.org/licenses/GPL-3.0 -->
   <!-- TODO: 
      * set up a simple batch file / shell file that will allow users to process a file directly in Saxon,
         without any other specialized software, e.g., Oxygen. -->


   <!-- = = = = = = = = = = = = = = = = = = = = = = = = = -->
   <!--         STUFF YOU CAN (SHOULD) CHANGE             -->
   <!-- = = = = = = = = = = = = = = = = = = = = = = = = = -->

   <!-- Where you do want to save the output? The default, tokenize(base-uri(/), '/')[last()] || '.docx',
      simply puts the Word file in the same directory, with the same name with the extension added. -->
   <xsl:param name="save-output-where" as="xs:string">{tokenize(base-uri(/), '/')[last()] ||
      '.docx'}</xsl:param>

   <!-- Where is the basic template you want to use? Perhaps you have your own specialized docx file. Cool.
      Point to it from here. Any relative reference will be resolved based upon the xml-to-docx.xsl application.
      If you want to use an absolute path, don't forget to use the "file:" protocol if it is local, e.g.,
      "file:/u:/mytemplate.docx". And yes, don't forget to use regular slashes, not backslashes (not that
      one is inherently superior to the other). -->
   <xsl:param name="template-docx-uri" as="xs:string">dependencies/template.docx</xsl:param>

   <!-- What other elements should begin a new paragraph? Generally speaking, if an element wraps non-space text, 
      it should occupy its own Word paragraph. But that element might have descendants that are paragraphs in 
      their own right. And you cannot nest paragraphs within Word. So the following parameter allows a special 
      work-around, allowing a user to declare elements that should prompt cutting that paragraph into multiple 
      paragraphs. For example, "^(list|item)$" says that any <list> or <item> should begin their own new 
      paragraph. The string value will be treated as a regular expression, using standard XPath regex rules.
   -->
   <xsl:param name="names-of-elements-that-should-begin-a-new-paragraph-regex" as="xs:string"
      >^(list|item)$</xsl:param>

   <!-- What type of space normalization do you want on text nodes? Expected is an integer:
      (1) moderate space normalization (default; all clumps of space characters are reduced to a single space 
         U+0020)
      (2) full space normalization (opening and closing space characters are deleted altogether, and then
         remaining clumps of space characters are reduced to a single space, U+0020)
      (3) no space normalization (all spaces are retained as-is)
      If you are working with mixed content (non-space text nodes interleave with elements that have non-space text
         nodes), you will likely prefer #1. A text node governed by @xml:space="preserve" will have #3 applied.
   -->
   <xsl:param name="space-normalization-level" as="xs:integer" select="3"/>

   <!-- What is the default point size for the main text in the template docx file? -->
   <xsl:param name="default-pt-size" as="xs:integer" select="24"/>

   <!-- Why in the world would you want the XML apparatus to be anything but smaller than the
      default docx point size? Here you can specify the default point size of opening and cloing 
      tags, comments, and so forth. But look out: maybe they have no effect. You should read the
      rest of this file. -->
   <xsl:param name="default-xml-apparatus-pt-size" as="xs:integer" select="12"/>


   <!-- Colors, colors, colors. That's where it's all at, to make sure that some things pop off
      the screen, and others stay subdued. These are all hexadecimal values from black, FFFFFF, 
      to white, 000000. -->

   <!-- What is the default font name of the special XML features? -->
   <xsl:param name="default-font-for-xml-apparatus" as="xs:string" select="'Courier New'"/>

   <!-- What color should opening and closing tags take? -->
   <xsl:param name="tag-color" as="xs:string" select="'000096'"/>
   <!-- What color should attribute names take? -->
   <xsl:param name="attr-name-color" as="xs:string" select="'F5844C'"/>
   <!-- There's an equals sign between an attribute name and its value. What color should it be? -->
   <xsl:param name="attr-equals-sign-color" as="xs:string" select="'FF8040'"/>
   <!-- What color should an attribute's value be? -->
   <xsl:param name="attr-value-color" as="xs:string" select="'993300'"/>
   <!-- What color should a comment be? -->
   <xsl:param name="comment-color" as="xs:string" select="'006400'"/>
   <!-- What color should an entity be? -->
   <xsl:param name="entity-color" as="xs:string" select="'969600'"/>
   <!-- What color should a processing instruction be? -->
   <xsl:param name="processing-instruction-color" as="xs:string" select="'8B26C9'"/>


   <!-- What should the default highlight value be for XML features? -->
   <xsl:param name="default-highlight-color" as="xs:string" select="'white'"/>

   <!-- For transparency settings below, any number from 0 to 99 will have effect; other values will
      be ignored. Settings: 0 has no transparency; 100 is completely transparent. -->

   <!-- What transparency value should opening and closing tags take? -->
   <xsl:param name="tag-transparency" as="xs:integer" select="5"/>
   <!-- What transparency value should attributes and their values take? -->
   <xsl:param name="attr-transparency" as="xs:integer" select="5"/>
   <!-- What transparency value should comments take? -->
   <xsl:param name="comment-transparency" as="xs:integer" select="5"/>
   <!-- What transparency value should entities take? -->
   <xsl:param name="entity-transparency" as="xs:integer" select="5"/>
   <!-- What transparency value should processing instructions take? -->
   <xsl:param name="processing-instruction-transparency" as="xs:integer" select="5"/>

   <!-- When the target docx is being built and saved, do you want messages during the process? If 
      true, you will be informed of the output path and which components are being packed into the 
      archive. This parameter belongs to the Open and Save Archive package. -->
   <xsl:param name="comment-on-saved-archives" as="xs:boolean" select="false()"/>


   <!-- = = = = = = = = = = = = = = = = = = = = = = = = = = = -->
   <!--  STUFF YOU SHOULD CHANGE ONLY IF YOU UNDERSTAND DOCX  -->
   <!--                  (you've been warned)                 -->
   <!-- = = = = = = = = = = = = = = = = = = = = = = = = = = = -->

   <xsl:param name="rPr-for-tag" as="element()">
      <w:rPr>
         <w:rFonts w:ascii="{$default-font-for-xml-apparatus}"
            w:hAnsi="{$default-font-for-xml-apparatus}" w:cs="{$default-font-for-xml-apparatus}"/>
         <w:color w:val="{$tag-color}"/>
         <w:sz w:val="{$default-xml-apparatus-pt-size}"/>
         <w:szCs w:val="{$default-xml-apparatus-pt-size}"/>
         <w:highlight w:val="{$default-highlight-color}"/>
         <xsl:if test="$tag-transparency ge 0 and $tag-transparency lt 100">
            <w14:textFill>
               <w14:solidFill>
                  <w14:srgbClr w14:val="{$tag-color}">
                     <w14:alpha w14:val="{$tag-transparency}000"/>
                  </w14:srgbClr>
               </w14:solidFill>
            </w14:textFill>
         </xsl:if>
      </w:rPr>
   </xsl:param>
   <xsl:param name="rPr-for-attr-name" as="element()">
      <w:rPr>
         <w:rFonts w:ascii="{$default-font-for-xml-apparatus}"
            w:hAnsi="{$default-font-for-xml-apparatus}" w:cs="{$default-font-for-xml-apparatus}"/>
         <w:color w:val="{$attr-name-color}"/>
         <w:sz w:val="{$default-xml-apparatus-pt-size}"/>
         <w:szCs w:val="{$default-xml-apparatus-pt-size}"/>
         <w:highlight w:val="{$default-highlight-color}"/>
         <xsl:if test="$attr-transparency ge 0 and $attr-transparency lt 100">
            <w14:textFill>
               <w14:solidFill>
                  <w14:srgbClr w14:val="{$attr-name-color}">
                     <w14:alpha w14:val="{$attr-transparency}000"/>
                  </w14:srgbClr>
               </w14:solidFill>
            </w14:textFill>
         </xsl:if>
      </w:rPr>
   </xsl:param>
   <xsl:param name="rPr-for-attr-equal-sign" as="element()">
      <w:rPr>
         <w:rFonts w:ascii="{$default-font-for-xml-apparatus}"
            w:hAnsi="{$default-font-for-xml-apparatus}" w:cs="{$default-font-for-xml-apparatus}"/>
         <w:color w:val="{$attr-equals-sign-color}"/>
         <w:sz w:val="{$default-xml-apparatus-pt-size}"/>
         <w:szCs w:val="{$default-xml-apparatus-pt-size}"/>
         <w:highlight w:val="{$default-highlight-color}"/>
         <xsl:if test="$attr-transparency ge 0 and $attr-transparency lt 100">
            <w14:textFill>
               <w14:solidFill>
                  <w14:srgbClr w14:val="{$attr-equals-sign-color}">
                     <w14:alpha w14:val="{$attr-transparency}000"/>
                  </w14:srgbClr>
               </w14:solidFill>
            </w14:textFill>
         </xsl:if>
      </w:rPr>
   </xsl:param>
   <xsl:param name="rPr-for-attr-value" as="element()">
      <w:rPr>
         <w:rFonts w:ascii="{$default-font-for-xml-apparatus}"
            w:hAnsi="{$default-font-for-xml-apparatus}" w:cs="{$default-font-for-xml-apparatus}"/>
         <w:color w:val="{$attr-value-color}"/>
         <w:sz w:val="{$default-xml-apparatus-pt-size}"/>
         <w:szCs w:val="{$default-xml-apparatus-pt-size}"/>
         <w:highlight w:val="{$default-highlight-color}"/>
         <xsl:if test="$attr-transparency ge 0 and $attr-transparency lt 100">
            <w14:textFill>
               <w14:solidFill>
                  <w14:srgbClr w14:val="{$attr-value-color}">
                     <w14:alpha w14:val="{$attr-transparency}000"/>
                  </w14:srgbClr>
               </w14:solidFill>
            </w14:textFill>
         </xsl:if>
      </w:rPr>
   </xsl:param>
   <xsl:param name="rPr-for-comment" as="element()">
      <w:rPr>
         <w:rFonts w:ascii="{$default-font-for-xml-apparatus}"
            w:hAnsi="{$default-font-for-xml-apparatus}" w:cs="{$default-font-for-xml-apparatus}"/>
         <w:color w:val="{$comment-color}"/>
         <w:sz w:val="{$default-xml-apparatus-pt-size}"/>
         <w:szCs w:val="{$default-xml-apparatus-pt-size}"/>
         <w:highlight w:val="{$default-highlight-color}"/>
         <xsl:if test="$comment-transparency ge 0 and $comment-transparency lt 100">
            <w14:textFill>
               <w14:solidFill>
                  <w14:srgbClr w14:val="{$comment-color}">
                     <w14:alpha w14:val="{$comment-transparency}000"/>
                  </w14:srgbClr>
               </w14:solidFill>
            </w14:textFill>
         </xsl:if>
      </w:rPr>
   </xsl:param>
   <xsl:param name="rPr-for-entity" as="element()">
      <w:rPr>
         <w:rFonts w:ascii="{$default-font-for-xml-apparatus}"
            w:hAnsi="{$default-font-for-xml-apparatus}" w:cs="{$default-font-for-xml-apparatus}"/>
         <w:color w:val="{$entity-color}"/>
         <w:sz w:val="{$default-pt-size}"/>
         <w:szCs w:val="{$default-pt-size}"/>
         <w:highlight w:val="{$default-highlight-color}"/>
         <xsl:if test="$entity-transparency ge 0 and $entity-transparency lt 100">
            <w14:textFill>
               <w14:solidFill>
                  <w14:srgbClr w14:val="{$entity-color}">
                     <w14:alpha w14:val="{$entity-transparency}000"/>
                  </w14:srgbClr>
               </w14:solidFill>
            </w14:textFill>
         </xsl:if>
      </w:rPr>
   </xsl:param>
   <xsl:param name="rPr-for-processing-instruction" as="element()">
      <w:rPr>
         <w:rFonts w:ascii="{$default-font-for-xml-apparatus}"
            w:hAnsi="{$default-font-for-xml-apparatus}" w:cs="{$default-font-for-xml-apparatus}"/>
         <w:color w:val="{$processing-instruction-color}"/>
         <w:sz w:val="{$default-xml-apparatus-pt-size}"/>
         <w:szCs w:val="{$default-xml-apparatus-pt-size}"/>
         <w:highlight w:val="{$default-highlight-color}"/>
         <xsl:if
            test="$processing-instruction-transparency ge 0 and $processing-instruction-transparency lt 100">
            <w14:textFill>
               <w14:solidFill>
                  <w14:srgbClr w14:val="{$processing-instruction-color}">
                     <w14:alpha w14:val="{$processing-instruction-transparency}000"/>
                  </w14:srgbClr>
               </w14:solidFill>
            </w14:textFill>
         </xsl:if>
      </w:rPr>
   </xsl:param>


   <!-- = = = = = = = = = = = = = = = = = = = = = = = = = = = -->
   <!--   STUFF YOU CAN CHANGE IF YOU KNOW WHAT YOU'RE DOING  -->
   <!--                  (you've been warned)                 -->
   <!-- = = = = = = = = = = = = = = = = = = = = = = = = = = = -->

   <xsl:variable name="primary-input" as="document-node()" select="."/>
   <xsl:variable name="template-docx" as="document-node()*"
      select="tan:open-docx(resolve-uri($template-docx-uri, static-base-uri()))"/>

   <xsl:variable name="output-pass-1" as="document-node()*">
      <xsl:apply-templates select="$template-docx" mode="output-pass-1"/>
   </xsl:variable>

   <xsl:mode name="output-pass-1" on-no-match="shallow-copy"/>

   <xsl:template match="w:document/w:body/w:p" priority="-1" mode="output-pass-1"/>
   <xsl:template match="w:document/w:body/w:p[1]" mode="output-pass-1">
      <xsl:variable name="standard-p" as="element()" select="."/>
      <xsl:apply-templates select="$primary-input" mode="xml-to-docx">
         <xsl:with-param name="standard-p" tunnel="yes" select="$standard-p"/>
      </xsl:apply-templates>
   </xsl:template>

   <xsl:mode name="xml-to-docx" on-no-match="shallow-skip"/>

   <xsl:template match="*" mode="xml-to-docx" priority="1">
      <xsl:param name="p-has-been-built" as="xs:boolean" tunnel="yes" select="false()"/>
      <xsl:param name="standard-p" as="element()" tunnel="yes"/>
      <xsl:variable name="wraps-nonspace-text" as="xs:boolean"
         select="exists(text()[matches(., '\S')])"/>
      <xsl:variable name="self-should-be-wrapped-in-p" as="xs:boolean"
         select="not($p-has-been-built) and $wraps-nonspace-text"/>
      <xsl:choose>
         <xsl:when test="$self-should-be-wrapped-in-p">
            <w:p>
               <xsl:copy-of select="$standard-p/@*"/>
               <xsl:apply-templates select="$standard-p/w:pPr" mode="adjust-pPr">
                  <xsl:with-param name="level" as="xs:integer" tunnel="yes"
                     select="count(ancestor::*)"/>
               </xsl:apply-templates>
               <xsl:next-match>
                  <xsl:with-param name="p-has-been-built" tunnel="yes" select="true()"/>
               </xsl:next-match>
            </w:p>
         </xsl:when>
         <xsl:otherwise>
            <xsl:next-match/>
         </xsl:otherwise>
      </xsl:choose>
   </xsl:template>

   <xsl:template match="*" mode="xml-to-docx">
      <xsl:param name="p-has-been-built" as="xs:boolean" tunnel="yes" select="false()"/>
      <xsl:param name="standard-p" as="element()" tunnel="yes"/>
      <xsl:param name="inherited-namespace-uri" tunnel="yes" as="xs:anyURI?" select="xs:anyURI('')"/>
      <xsl:param name="level" tunnel="yes" as="xs:integer" select="count(ancestor::*)"/>
      <xsl:param name="secondary-level" tunnel="yes" as="xs:integer" select="count(ancestor::*) + 1"/>

      <xsl:variable name="this-default-namespace" as="xs:anyURI?"
         select="namespace-uri-for-prefix('', .)"/>
      <xsl:variable name="element-name" as="xs:string" select="name(.)"/>
      <xsl:variable name="this-element-should-break" as="xs:boolean"
         select="matches($element-name, $names-of-elements-that-should-begin-a-new-paragraph-regex)"/>
      <xsl:variable name="wraps-nonspace-text" as="xs:boolean"
         select="exists(text()[matches(., '\S')])"/>
      <xsl:variable name="wraps-mixed-content" as="xs:boolean"
         select="exists(*) and $wraps-nonspace-text"/>
      <xsl:variable name="docx-element-to-pass-through" as="element()" select="
            if ($p-has-been-built) then
               $standard-p/w:r[1]
            else
               $standard-p"/>

      <!-- If the element deserves its own paragraph (a pseudo-paragraph within a paragraph), prime it
         for the next pass. -->
      <xsl:if test="$p-has-been-built and $this-element-should-break">
         <break indent="{$secondary-level}"/>
      </xsl:if>
      <xsl:apply-templates select="$docx-element-to-pass-through" mode="create-tag-in-docx">
         <xsl:with-param name="new-content" tunnel="yes" select="."/>
         <xsl:with-param name="is-opening-tag" as="xs:boolean" tunnel="yes" select="true()"/>
         <xsl:with-param name="level" as="xs:integer" tunnel="yes" select="count(ancestor::*)"/>
      </xsl:apply-templates>
      <xsl:choose>
         <xsl:when test="$wraps-nonspace-text and not($p-has-been-built)">
            <w:p>
               <xsl:apply-templates select="$standard-p/w:pPr" mode="adjust-pPr">
                  <xsl:with-param name="level" as="xs:integer" tunnel="yes"
                     select="count(ancestor::*)"/>
               </xsl:apply-templates>
               <xsl:apply-templates mode="#current">
                  <xsl:with-param name="p-has-been-built" tunnel="yes" select="true()"/>
                  <xsl:with-param name="inherited-namespace-uri" tunnel="yes"
                     select="$this-default-namespace"/>
               </xsl:apply-templates>
            </w:p>
         </xsl:when>
         <xsl:otherwise>
            <xsl:apply-templates mode="#current">
               <xsl:with-param name="inherited-namespace-uri" tunnel="yes"
                  select="$this-default-namespace"/>
               <xsl:with-param name="secondary-level" tunnel="yes" as="xs:integer"
                  select="$secondary-level + 1"/>
            </xsl:apply-templates>

         </xsl:otherwise>
      </xsl:choose>
      <!-- Make a closing tag only if needed -->
      <xsl:if test="exists(node())">
         <xsl:apply-templates select="$docx-element-to-pass-through" mode="create-tag-in-docx">
            <xsl:with-param name="new-content" tunnel="yes" select="."/>
            <xsl:with-param name="is-opening-tag" as="xs:boolean" tunnel="yes" select="false()"/>
            <xsl:with-param name="level" as="xs:integer" tunnel="yes" select="count(ancestor::*)"/>
         </xsl:apply-templates>
      </xsl:if>
      <!-- Prime the closing of the pseudo-paragraph -->
      <xsl:if test="$p-has-been-built and $this-element-should-break">
         <break indent="{$secondary-level - 1}"/>
      </xsl:if>
   </xsl:template>

   <xsl:template match="comment() | processing-instruction()" mode="xml-to-docx" priority="1">
      <xsl:param name="standard-p" as="element()" tunnel="yes"/>
      <xsl:param name="p-has-been-built" as="xs:boolean" tunnel="yes" select="false()"/>
      <xsl:choose>
         <xsl:when test="$p-has-been-built">
            <xsl:next-match/>
         </xsl:when>
         <xsl:otherwise>
            <w:p>
               <xsl:copy-of select="$standard-p/@*"/>
               <xsl:apply-templates select="$standard-p/w:pPr" mode="adjust-pPr">
                  <xsl:with-param name="level" as="xs:integer" tunnel="yes"
                     select="count(ancestor::*)"/>
               </xsl:apply-templates>
               <xsl:next-match/>
            </w:p>
         </xsl:otherwise>
      </xsl:choose>
   </xsl:template>

   <xsl:template match="comment()" mode="xml-to-docx" expand-text="yes">
      <xsl:param name="secondary-level" tunnel="yes" as="xs:integer" select="count(ancestor::*) + 1"/>

      <xsl:variable name="comment-lines" as="xs:string+" select="tokenize(., '\r?\n')"/>
      <xsl:variable name="comment-line-count" as="xs:integer" select="count($comment-lines)"/>
      <xsl:for-each select="$comment-lines">
         <xsl:if test="position() gt 1">
            <break indent="{$secondary-level}"/>
         </xsl:if>
         <w:r>
            <xsl:copy-of select="$rPr-for-comment"/>
            <xsl:choose>
               <xsl:when test="$comment-line-count eq 1">
                  <w:t xml:space="preserve">&lt;!--{.}--&gt;</w:t>
               </xsl:when>
               <xsl:when test="position() eq 1">
                  <w:t xml:space="preserve">&lt;!--{.}</w:t>
               </xsl:when>
               <xsl:when test="position() eq $comment-line-count">
                  <w:t xml:space="preserve">{.}--&gt;</w:t>
               </xsl:when>
               <xsl:otherwise>
                  <w:t xml:space="preserve">{.}</w:t>
               </xsl:otherwise>
            </xsl:choose>
         </w:r>
      </xsl:for-each>
   </xsl:template>

   <xsl:template match="processing-instruction()" mode="xml-to-docx" expand-text="yes">
      <xsl:param name="secondary-level" tunnel="yes" as="xs:integer" select="count(ancestor::*) + 1"/>

      <xsl:variable name="pi-lines" as="xs:string+" select="tokenize(., '\r?\n')"/>
      <xsl:variable name="pi-line-count" as="xs:integer" select="count($pi-lines)"/>
      <xsl:variable name="pi-name" as="xs:string" select="name(.)"/>
      <xsl:for-each select="$pi-lines">
         <xsl:if test="position() gt 1">
            <break indent="{$secondary-level}"/>
         </xsl:if>
         <w:r>
            <xsl:copy-of select="$rPr-for-processing-instruction"/>
            <xsl:choose>
               <xsl:when test="$pi-line-count eq 1">
                  <w:t xml:space="preserve">&lt;?{$pi-name} {.}?&gt;</w:t>
               </xsl:when>
               <xsl:when test="position() eq 1">
                  <w:t xml:space="preserve">&lt;?{$pi-name} {.}</w:t>
               </xsl:when>
               <xsl:when test="position() eq $pi-line-count">
                  <w:t xml:space="preserve">{.}?&gt;</w:t>
               </xsl:when>
               <xsl:otherwise>
                  <w:t xml:space="preserve">{.}</w:t>
               </xsl:otherwise>
            </xsl:choose>
         </w:r>
      </xsl:for-each>
   </xsl:template>

   <xsl:template match="text()" mode="xml-to-docx">
      <xsl:param name="standard-p" as="element()" tunnel="yes"/>
      <xsl:param name="p-has-been-built" as="xs:boolean" tunnel="yes" select="false()"/>
      <xsl:param name="secondary-level" tunnel="yes" as="xs:integer" select="count(ancestor::*) + 1"/>

      <xsl:variable name="keep-space-as-is" as="xs:boolean" select="
            $space-normalization-level eq 3
            or (ancestor-or-self::*[@xml:space][1]/@xml:space = 'preserve')"/>
      <xsl:variable name="drop-end-or-closing-space" as="xs:boolean" select="
            not($keep-space-as-is)
            and $space-normalization-level eq 2"/>
      <xsl:variable name="this-text-norm" as="xs:string*">
         <xsl:choose>
            <xsl:when test="$keep-space-as-is">
               <xsl:value-of select="."/>
            </xsl:when>
            <xsl:otherwise>
               <xsl:if test="not($drop-end-or-closing-space) and matches(., '^\s')">
                  <xsl:value-of select="' '"/>
               </xsl:if>
               <xsl:value-of select="normalize-space(.)"/>
               <xsl:if
                  test="not($drop-end-or-closing-space) and matches(., '\s$') and matches(., '\S')">
                  <xsl:value-of select="' '"/>
               </xsl:if>
            </xsl:otherwise>
         </xsl:choose>
      </xsl:variable>
      <xsl:variable name="these-rs" as="element()*">
         <xsl:for-each select="tokenize(string-join($this-text-norm), '\r?\n')">
            <xsl:if test="position() gt 1">
               <break indent="0"/>
            </xsl:if>
            <xsl:analyze-string select="." regex="[&amp;&lt;]">
               <xsl:matching-substring>
                  <w:r>
                     <xsl:copy-of select="$rPr-for-entity"/>
                     <w:t>
                        <xsl:attribute name="xml:space" select="'preserve'"/>
                        <xsl:value-of select="if (. eq '&amp;') then '&amp;amp;' else '&amp;lt;'"/>
                     </w:t>
                  </w:r>
               </xsl:matching-substring>
               <xsl:non-matching-substring>
                  <w:r>
                     <w:t>
                        <xsl:attribute name="xml:space" select="'preserve'"/>
                        <xsl:value-of select="."/>
                     </w:t>
                  </w:r>
               </xsl:non-matching-substring>
            </xsl:analyze-string>

         </xsl:for-each>



      </xsl:variable>
      <xsl:choose>
         <xsl:when test="$p-has-been-built">
            <xsl:copy-of select="$these-rs"/>
         </xsl:when>
         <xsl:when test="matches(., '\S')">
            <w:p>
               <xsl:copy-of select="$standard-p/@*"/>
               <xsl:apply-templates select="$standard-p/w:pPr" mode="adjust-pPr">
                  <xsl:with-param name="level" as="xs:integer" tunnel="yes"
                     select="count(ancestor::*)"/>
               </xsl:apply-templates>
               <xsl:copy-of select="$these-rs"/>
            </w:p>
         </xsl:when>
      </xsl:choose>
   </xsl:template>

   <xsl:mode name="create-tag-in-docx" on-no-match="shallow-copy"/>

   <xsl:template match="w:p/w:r" priority="-1" mode="create-tag-in-docx"/>
   <xsl:template match="w:p/w:r[1]" mode="create-tag-in-docx" expand-text="yes">
      <xsl:param name="new-content" tunnel="yes" as="element()"/>
      <xsl:param name="is-opening-tag" tunnel="yes" as="xs:boolean" select="true()"/>
      <w:r>
         <xsl:copy-of select="$rPr-for-tag"/>
         <w:t>&lt;{if ($is-opening-tag) then () else '/'}{name($new-content)}</w:t>
      </w:r>
      <xsl:if test="$is-opening-tag">
         <xsl:apply-templates select="$new-content/namespace-node()" mode="attr-to-docx"/>
         <xsl:apply-templates select="$new-content/@*" mode="attr-to-docx"/>
      </xsl:if>
      <w:r>
         <xsl:copy-of select="$rPr-for-tag"/>
         <w:t>{if ($is-opening-tag and not(exists($new-content/node()))) then '/' else ()}&gt;</w:t>
      </w:r>
   </xsl:template>

   <xsl:mode name="attr-to-docx" on-no-match="shallow-skip"/>
   <xsl:template match="namespace-node()[name() eq 'xml']" priority="1" mode="attr-to-docx"/>
   <xsl:template match="namespace-node()" mode="attr-to-docx" expand-text="yes">
      <xsl:param name="inherited-namespace-uri" tunnel="yes" as="xs:anyURI?" select="xs:anyURI('')"/>

      <xsl:variable name="this-name" as="xs:string?" select="
            if (string-length(name(.)) lt 1) then
               'xmlns'
            else
               'xmlns:' || name(.)"/>
      <xsl:if test="string-length(name(.)) gt 0 or not(xs:anyURI(.) eq $inherited-namespace-uri)">
         <w:r>
            <xsl:copy-of select="$rPr-for-attr-name"/>
            <w:t xml:space="preserve"> {$this-name}</w:t>
         </w:r>
         <w:r>
            <xsl:copy-of select="$rPr-for-attr-equal-sign"/>
            <w:t>=</w:t>
         </w:r>
         <w:r>
            <xsl:copy-of select="$rPr-for-attr-value"/>
            <w:t>"{.}"</w:t>
         </w:r>
      </xsl:if>
   </xsl:template>
   <xsl:template match="@*" mode="attr-to-docx" expand-text="yes">
      <w:r>
         <xsl:copy-of select="$rPr-for-attr-name"/>
         <w:t xml:space="preserve"> {name(.)}</w:t>
      </w:r>
      <w:r>
         <xsl:copy-of select="$rPr-for-attr-equal-sign"/>
         <w:t>=</w:t>
      </w:r>
      <w:r>
         <xsl:copy-of select="$rPr-for-attr-value"/>
         <w:t>"{.}"</w:t>
      </w:r>
   </xsl:template>

   <xsl:param name="indentation-factor" as="xs:integer" select="120"/>

   <xsl:mode name="adjust-pPr" on-no-match="shallow-copy"/>

   <xsl:template match="w:pPr/w:ind" mode="adjust-pPr create-tag-in-docx">
      <xsl:param name="level" tunnel="yes" as="xs:integer"/>
      <xsl:copy>
         <xsl:copy-of select="@*"/>
         <xsl:attribute name="w:left" select="$level * $indentation-factor"/>
      </xsl:copy>
   </xsl:template>
   <xsl:template match="w:pPr[not(w:ind)]/*[1]" mode="adjust-pPr create-tag-in-docx">
      <xsl:param name="level" tunnel="yes" as="xs:integer"/>
      <w:ind w:left="{$level * $indentation-factor}"/>
      <xsl:copy>
         <xsl:copy-of select="@*"/>
         <xsl:apply-templates mode="#current"/>
      </xsl:copy>
   </xsl:template>



   <!-- The strategy behind pass 2 is to replace each <break> with a new paragraph,
      and to make some slight adjustments where needed to indentation. -->

   <xsl:variable name="output-pass-2" as="document-node()*">
      <xsl:apply-templates select="$output-pass-1" mode="chop-paragraphs"/>
   </xsl:variable>

   <xsl:mode name="chop-paragraphs" on-no-match="shallow-copy"/>

   <xsl:template match="break" mode="chop-paragraphs"/>

   <xsl:template match="w:p[break]" mode="chop-paragraphs">
      <xsl:variable name="current-w-pPr" as="element()?" select="w:pPr"/>
      <xsl:variable name="this-p" as="element()" select="."/>

      <xsl:for-each-group select="node()" group-starting-with="break">
         <xsl:variable name="break-marker" as="element()?" select="current-group()[1]/self::break"/>
         <w:p>
            <xsl:copy-of select="$this-p/@*"/>
            <xsl:if test="exists($break-marker)">
               <xsl:apply-templates select="$current-w-pPr" mode="adjust-indentation">
                  <xsl:with-param name="indentation" tunnel="yes"
                     select="xs:integer($break-marker/@indent)"/>
               </xsl:apply-templates>
            </xsl:if>
            <xsl:apply-templates select="current-group()" mode="#current"/>
         </w:p>
      </xsl:for-each-group>

   </xsl:template>

   <xsl:mode name="adjust-indentation" on-no-match="shallow-copy"/>

   <xsl:template match="w:ind" mode="adjust-indentation"/>

   <xsl:template match="w:pPr" mode="adjust-indentation">
      <xsl:param name="indentation" tunnel="yes" as="xs:integer"/>
      <xsl:copy>
         <xsl:copy-of select="@*"/>
         <w:ind w:left="{$indentation * $indentation-factor}"/>
         <xsl:apply-templates mode="#current"/>
      </xsl:copy>
   </xsl:template>

   <!-- Do you want to output diagnostics within the primary output? Could be useful if you want to look 
      inside the black box. (I hate black boxes that you can't open. So this one has quick release hinges.) -->
   <xsl:param name="diagnostics-on" as="xs:boolean" static="yes" select="false()"/>
   <xsl:output indent="yes" omit-xml-declaration="yes"/>
   <xsl:template match="/" priority="1" use-when="$diagnostics-on">
      <diagnostics>
         <output-pass-1-document>
            <xsl:copy-of select="$output-pass-1[w:document]"/>
         </output-pass-1-document>
         <output-pass-2-document>
            <xsl:copy-of select="$output-pass-2[w:document]"/>
         </output-pass-2-document>
      </diagnostics>
      <xsl:next-match/>
   </xsl:template>

   <xsl:template match="/">
      <xsl:call-template name="tan:save-docx">
         <xsl:with-param name="docx-components" select="$output-pass-2"/>
         <xsl:with-param name="resolved-uri" select="resolve-uri($save-output-where, base-uri(/))"/>
      </xsl:call-template>
   </xsl:template>

   <xsl:import href="dependencies/open-and-save-docx.xsl"/>

</xsl:stylesheet>