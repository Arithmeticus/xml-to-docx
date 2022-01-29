<?xml version="1.0" encoding="UTF-8"?>
<sch:schema xmlns:sch="http://purl.oclc.org/dsdl/schematron" 
   xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
   queryBinding="xslt2"
   xmlns:sqf="http://www.schematron-quickfix.com/validator/process">
   <sch:title>Schematron tests for maintaining the TAN function library</sch:title>
   <sch:ns uri="http://www.w3.org/1999/XSL/Transform" prefix="xsl"/>
   <sch:ns uri="tag:textalign.net,2015:ns" prefix="tan"/>
   
   <sch:let name="master-open-and-save-archive-file-resolved-uri"
      value="'https://github.com/Arithmeticus/xslt-for-docx/raw/master/open-and-save-archive.xsl'"/>
   <sch:let name="master-open-and-save-archive-file"
      value="doc($master-open-and-save-archive-file-resolved-uri)"/>
   <sch:pattern>
      <sch:title>External inclusions</sch:title>
      <sch:p>These rules pertain to stylesheets that are part-and-parcel of external libraries. This
         approach has been adopted instead of Git submodule because of the difficulty for
         nonexperienced users to make sure submodules are included. </sch:p>
      <sch:rule
         context="xsl:stylesheet[contains(base-uri(.), 'open-and-save')]/* | 
         xsl:stylesheet[contains(base-uri(.), 'open-and-save')]/comment()">
         <sch:let name="current-path" value="path(.)"/>
         <sch:let name="corresponding-element"
            value="$master-open-and-save-archive-file/*/node()[path(.) eq $current-path]"/>
         <sch:assert role="warning" test="exists($corresponding-element)">There is no node in the master file that
            corresponds to this one, <sch:value-of select="$current-path"/>. Download the 
            latest at <sch:value-of select="$master-open-and-save-archive-file-resolved-uri"/></sch:assert>
         <sch:assert role="warning" test="exists($corresponding-element) and deep-equal(., $corresponding-element)"
            sqf:fix="copyMasterNode">This node does not match its counterpart. <sch:value-of select="serialize($corresponding-element)"/> Download the 
            latest at <sch:value-of select="$master-open-and-save-archive-file-resolved-uri"/></sch:assert>
         
         <sqf:fix id="copyMasterNode" use-when="exists($corresponding-element)">
            <sqf:description>
               <sqf:title>Replace current node with the corresponding master node</sqf:title>
            </sqf:description>
            <sqf:replace select="$corresponding-element"/>
         </sqf:fix>
      </sch:rule>
   </sch:pattern>
   
   
</sch:schema>