#!/usr/bin/env python3
"""
inject_slides.py - TOWING Presentation Integration Tool

This script implements the "Reverse Injection" pattern for generating robust PowerPoint files.
Instead of applying a master to a generated file (which causes corruption), it injects
generated content slides INTO a valid, pristine TOWING template.

Usage:
    python3 inject_slides.py template.pptx content.pptx output.pptx

Core Logic:
1. Unpack both Template and Content PPTX files.
2. Copy slide XML files from Content to Template.
3. Sanitize Slides:
   - Remove <p:bg> (background) to reveal Master background.
   - Remove <p:ph> (placeholders) to detach from mismatched layouts.
4. Normalize Layouts:
   - Rewrite _rels to force all slides to use Template's 'slideLayout1.xml'.
   - Remove references to non-existent files (like notesSlides).
5. Merge Metadata:
   - Renumber Relationship IDs (rId) to avoid collisions.
   - Update presentation.xml (Slide List) maintaining strict tag ordering.
   - Update [Content_Types].xml.
6. Repack into a new PPTX.
"""

import sys
import os
import shutil
import tempfile
import zipfile
import re
from pathlib import Path
import xml.etree.ElementTree as ET

def register_namespace_prefixes():
    """Register OOXML namespaces to prevent ns0: prefixes in output"""
    ET.register_namespace('', 'http://schemas.openxmlformats.org/presentationml/2006/main')
    ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    ET.register_namespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
    ET.register_namespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types')

def get_rel_max_id(root):
    """Find the highest rId (e.g., rId5) in a relationships file"""
    max_id = 0
    for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rid = rel.get('Id')
        if rid and rid.startswith('rId'):
            try:
                num = int(rid[3:])
                if num > max_id:
                    max_id = num
            except ValueError:
                pass
    return max_id

def inject_slides(template_pptx, content_pptx, output_pptx):
    print(f"Injecting slides from {content_pptx} into {template_pptx}...")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        tpl_dir = temp_path / "template"
        cnt_dir = temp_path / "content"
        
        # Unzip both presentations
        with zipfile.ZipFile(template_pptx, 'r') as z: z.extractall(tpl_dir)
        with zipfile.ZipFile(content_pptx, 'r') as z: z.extractall(cnt_dir)
        
        # --- Step 1: Copy Slide XMLs ---
        src_slides_dir = cnt_dir / 'ppt' / 'slides'
        dest_slides_dir = tpl_dir / 'ppt' / 'slides'
        dest_slides_dir.mkdir(parents=True, exist_ok=True)
        
        src_slides = sorted(list(src_slides_dir.glob('slide*.xml')))
        print(f"  Found {len(src_slides)} source slides.")
        
        # Directory for slide relationships
        src_slides_rels = cnt_dir / 'ppt' / 'slides' / '_rels'
        dest_slides_rels = tpl_dir / 'ppt' / 'slides' / '_rels'
        dest_slides_rels.mkdir(parents=True, exist_ok=True)
        
        for slide in src_slides:
            # Copy slide XML file
            shutil.copy2(slide, dest_slides_dir / slide.name)
            
            # Clean slide XML (Crucial for stability)
            try:
                with open(dest_slides_dir / slide.name, 'r', encoding='utf-8') as f:
                    xml = f.read()
                
                # Remove background (<p:bg>) to let Master background show through
                xml = re.sub(r'<p:bg>.*?</p:bg>', '', xml, flags=re.DOTALL)
                
                # Remove placeholders (<p:ph>) to prevent layout mismatch errors
                # This detaches the content from any specific layout logic
                xml = re.sub(r'<p:ph\s+[^>]*?>', '', xml)
                
                with open(dest_slides_dir / slide.name, 'w', encoding='utf-8') as f:
                    f.write(xml)
            except Exception as e:
                print(f"    Warning cleaning {slide.name}: {e}")

            # Process Relationship Files (.rels)
            # We must ensure all targets exist.
            rel_name = slide.name + ".rels"
            if (src_slides_rels / rel_name).exists():
                tree = ET.parse(src_slides_rels / rel_name)
                root = tree.getroot()
                ns = {'': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                
                rels_to_remove = []
                for rel in root.findall('Relationship', ns):
                    type_val = rel.get('Type').split('/')[-1]
                    
                    if 'slideLayout' in type_val:
                        # FORCE Target to Template's standard layout (slideLayout1.xml)
                        # This normalizes all slides to use a known, valid layout
                        rel.set('Target', '../slideLayouts/slideLayout1.xml')
                    
                    elif 'notesSlide' in type_val:
                        # Remove notesSlide references since we DO NOT copy notes files
                        # Leaving this reference would cause file corruption
                        rels_to_remove.append(rel)
                
                for rel in rels_to_remove:
                    root.remove(rel)
                
                tree.write(dest_slides_rels / rel_name, encoding='UTF-8', xml_declaration=True)

        # --- Step 2: Merge Global Relationships (presentation.xml.rels) ---
        # We need to register the new slides in the Template's relationship file
        # AND we must renumber IDs to prevent collision.
        
        cnt_rels_tree = ET.parse(cnt_dir / 'ppt' / '_rels' / 'presentation.xml.rels')
        cnt_rels_root = cnt_rels_tree.getroot()
        
        tpl_rels_path = tpl_dir / 'ppt' / '_rels' / 'presentation.xml.rels'
        tpl_rels_tree = ET.parse(tpl_rels_path)
        tpl_rels_root = tpl_rels_tree.getroot()
        
        rns = {'': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        
        # Calculate max rId in Template to start numbering after it
        max_rid = get_rel_max_id(tpl_rels_root)
        print(f"  Template max rId: {max_rid}")
        
        # Map old rId (from Content) to new rId (in Template)
        rid_map = {}
        
        # Remove existing slide references from Template (we are replacing them)
        for rel in tpl_rels_root.findall('Relationship', rns):
            if 'slide' == rel.get('Type').split('/')[-1]:
                tpl_rels_root.remove(rel)
        
        # Add new slide relationships with NEW IDs
        current_rid = max_rid
        for rel in cnt_rels_root.findall('Relationship', rns):
            if 'slide' == rel.get('Type').split('/')[-1]:
                current_rid += 1
                old_rid = rel.get('Id')
                new_rid = f"rId{current_rid}"
                
                rel.set('Id', new_rid)
                tpl_rels_root.append(rel)
                
                rid_map[old_rid] = new_rid
                print(f"    Remapped slide {old_rid} -> {new_rid}")
        
        tpl_rels_tree.write(tpl_rels_path, encoding='UTF-8', xml_declaration=True)

        # --- Step 3: Update Presentation.xml (Slide List) ---
        register_namespace_prefixes()
        
        cnt_pres_tree = ET.parse(cnt_dir / 'ppt' / 'presentation.xml')
        cnt_pres_root = cnt_pres_tree.getroot()
        ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
              'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        
        cnt_sldIdLst = cnt_pres_root.find('p:sldIdLst', ns)
        
        # Update r:id in the slide list using our map
        if cnt_sldIdLst is not None:
            for sldId in cnt_sldIdLst.findall('p:sldId', ns):
                old_rid = sldId.get(f"{{{ns['r']}}}id")
                if old_rid in rid_map:
                    sldId.set(f"{{{ns['r']}}}id", rid_map[old_rid])
        
        # Load Template presentation.xml
        tpl_pres_path = tpl_dir / 'ppt' / 'presentation.xml'
        tpl_pres_tree = ET.parse(tpl_pres_path)
        tpl_pres_root = tpl_pres_tree.getroot()
        
        # Replace sldIdLst
        # IMPORTANT: We must preserve XML tag order. sldIdLst cannot be just appended.
        tpl_sldIdLst = tpl_pres_root.find('p:sldIdLst', ns)
        
        insert_index = -1
        if tpl_sldIdLst is not None:
            # Find index of existing list
            for i, child in enumerate(tpl_pres_root):
                if child == tpl_sldIdLst:
                    insert_index = i
                    break
            tpl_pres_root.remove(tpl_sldIdLst)
        else:
            # Try to insert after master list if sldIdLst doesn't exist
            master_lst = tpl_pres_root.find('p:sldMasterIdLst', ns)
            if master_lst is not None:
                for i, child in enumerate(tpl_pres_root):
                    if child == master_lst:
                        insert_index = i + 1
                        break
        
        # Insert at correct position or append if fallback
        if cnt_sldIdLst is not None:
            if insert_index != -1:
                tpl_pres_root.insert(insert_index, cnt_sldIdLst)
            else:
                tpl_pres_root.append(cnt_sldIdLst)
        
        tpl_pres_tree.write(tpl_pres_path, encoding='UTF-8', xml_declaration=True)

        # --- Step 4: Update [Content_Types].xml ---
        cnt_ct_tree = ET.parse(cnt_dir / '[Content_Types].xml')
        cnt_ct_root = cnt_ct_tree.getroot()
        
        tpl_ct_path = tpl_dir / '[Content_Types].xml'
        tpl_ct_tree = ET.parse(tpl_ct_path)
        tpl_ct_root = tpl_ct_tree.getroot()
        
        cns = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
        
        # Copy 'Override' elements for slides ONLY
        existing_parts = {o.get('PartName') for o in tpl_ct_root.findall('ct:Override', cns)}
        
        for override in cnt_ct_root.findall('ct:Override', cns):
            pn = override.get('PartName')
            # Only add slides. Skip layouts and notes.
            if '/ppt/slides/' in pn and pn not in existing_parts:
                tpl_ct_root.append(override)
                
        tpl_ct_tree.write(tpl_ct_path, encoding='UTF-8', xml_declaration=True)

        # --- Step 5: Pack ---
        with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as z:
            for file_path in tpl_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(tpl_dir)
                    z.write(file_path, arcname)

    print(f"✅ Slides injected successfully: {output_pptx}")

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python3 inject_slides.py template.pptx content.pptx output.pptx")
        sys.exit(1)
    
    inject_slides(sys.argv[1], sys.argv[2], sys.argv[3])