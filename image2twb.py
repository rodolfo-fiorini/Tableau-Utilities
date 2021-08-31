# Copyright 2021 Apogee Integration, LLC

# Licensed under the Apache License, Version 2.0 (the "License"); # you may not use this file except in compliance with the License.
# You may obtain a copy of the License at

#   https://urldefense.us/v3/__http://www.apache.org/licenses/LICENSE-2.0__;!!BClRuOV5cvtbuNI!XHSjK-altBMEBcb29doVVSN5b_-KMFT0XOr_oxG_8N6hXeQc0NUoTQmpgimnDkMRItPthShdnydfyoZZ$ 

# Unless required by applicable law or agreed to in writing, software # distributed under the License is distributed on an "AS IS" BASIS, # WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and # limitations under the License.

import os
import argparse
import subprocess

# Returns a list of sorted image files within directory ("path") base on their time stamp of creation
def image_filenames(path):
  image_extensions = ('jpg','jpeg', 'png', 'bmp')
  file_names = [fn for fn in os.listdir(path) if fn.lower().endswith(image_extensions)]
  return list(sorted(file_names, key=lambda f: os.stat(os.path.join(path, f)).st_mtime))

def write_dashboard(file, size, image_path, num):
  file.write("""

    <dashboard name='Dashboard {slide_num}'>
      <style />
      <size {sizing_string} />
      <zones>
        <zone h='100000' id='4' type='layout-basic' w='100000' x='0' y='0'>
          <zone h='97090' id='3' is-centered='0' is-scaled='1' param='{image_folder_image_path}' type='bitmap' w='98316' x='842' y='1455'>
            <zone-style>
              <format attr='border-color' value='#000000' />
              <format attr='border-style' value='none' />
              <format attr='border-width' value='0' />
              <format attr='margin' value='4' />
            </zone-style>
          </zone>
          <zone-style>
            <format attr='border-color' value='#000000' />
            <format attr='border-style' value='none' />
            <format attr='border-width' value='0' />
            <format attr='margin' value='8' />
          </zone-style>
        </zone>
      </zones>
      <devicelayouts>
        <devicelayout auto-generated='true' name='Phone'>
          <size maxheight='700' minheight='700' sizing-mode='vscroll' />
          <zones>
            <zone h='100000' id='6' type='layout-basic' w='100000' x='0' y='0'>
              <zone h='97090' id='5' param='vert' type='layout-flow' w='98316' x='842' y='1455'>
                <zone fixed-size='280' h='97090' id='3' is-centered='0' is-fixed='true' is-scaled='1' param='{image_folder_image_path}' type='bitmap' w='98316' x='842' y='1455'>
                  <zone-style>
                    <format attr='border-color' value='#000000' />
                    <format attr='border-style' value='none' />
                    <format attr='border-width' value='0' />
                    <format attr='margin' value='4' />
                    <format attr='padding' value='0' />
                  </zone-style>
                </zone>
              </zone>
              <zone-style>
                <format attr='border-color' value='#000000' />
                <format attr='border-style' value='none' />
                <format attr='border-width' value='0' />
                <format attr='margin' value='8' />
              </zone-style>
            </zone>
          </zones>
        </devicelayout>
      </devicelayouts>
      <simple-id uuid='{{4D058E49-AB62-4056-BA04-B1F1036B{end_id}}}' />
    </dashboard>""".format(image_folder_image_path = image_path , slide_num = num, end_id = num + 1000, sizing_string = size))

def write_story_point(file, num):
  file.write("""
    <window class='dashboard' hidden='true' maximized='true' name='Dashboard {slide_num}'>
      <viewpoints />
      <active id='-1' />
      <simple-id uuid='{{B37FC551-7DBC-47F4-8E07-908C28F9{end_id}}}' />
    </window>""".format(slide_num = num, end_id = num + 1000))
 
def create_story_workbook(args):
 
  if args.fixed is None:
    sizing_string = "sizing-mode='automatic'"
    story_sizing_string = "sizing-mode='automatic'"
  else:
    height = args.fixed[0][0]
    width = args.fixed[0][1]
    sizing_string = "maxheight='{height}' maxwidth='{width}' minheight='{height}' minwidth='{width}' sizing-mode='fixed'".format(height = height, width = width)
    story_sizing_string = "maxheight='964' maxwidth='1016' minheight='964' minwidth='1016'"

  if args.tableau_path_name.endswith(".twb"):
    tableau_file_path = args.tableau_path_name
  else:
    tableau_file_path = args.tableau_path_name + ".twb"

  if os.path.exists(tableau_file_path) and not args.replace:
    print("File {path} already exists. If you wish to replace an existing file, use the -r or --replace flag.".format(path = tableau_file_path))
    exit(0)
 
  if not os.path.exists(args.images_folder_path):
    print("No folder named {dir}".format(dir = args.images_folder_path))
    exit(0)

  directory = os.path.abspath(args.images_folder_path)
  image_list = image_filenames(directory)
  
  if not image_list:
    print("Folder {dir} does not contain any image files.".format(dir = args.images_folder_path))
    exit(0)

  with open(tableau_file_path, 'w') as f:
    f.write(

  """<?xml version='1.0' encoding='utf-8' ?>
  <!-- build 20194.19.1010.1202                               -->
  <workbook original-version='18.1' source-build='2019.4.0 (20194.19.1010.1202)' source-platform='mac' version='18.1' xmlns:user='http://www.tableausoftware.com/xml/user'>
  <document-format-change-manifest>
    <AutoCreateAndUpdateDSDPhoneLayouts ignorable='true' predowngraded='true' />
    <SheetIdentifierTracking ignorable='true' predowngraded='true' />
    <WindowsPersistSimpleIdentifiers />
  </document-format-change-manifest>
  <preferences>
    <preference name='ui.encoding.shelf.height' value='24' />
    <preference name='ui.shelf.height' value='26' />
  </preferences>
  <datasources />
  <dashboards>""")

    dashboard_num = 1
    for image_file in image_list:
      write_dashboard(f, sizing_string, os.path.join(directory, image_file), dashboard_num)
      dashboard_num += 1

    # Story header
    f.write("""
    <dashboard name='Story 1' type='storyboard'>
      <style />
      <size {story_sizing_string}/>
      <zones>
        <zone h='100000' id='2' type='layout-basic' w='100000' x='0' y='0'>
          <zone h='98340' id='1' param='vert' removable='false' type='layout-flow' w='98426' x='787' y='830'>
            <zone h='3423' id='3' type='title' w='98426' x='787' y='830' />
            <zone h='10477' id='4' is-fixed='true' paired-zone-id='5' removable='false' type='flipboard-nav' w='98426' x='787' y='4253' />
            <zone h='84440' id='5' paired-zone-id='4' removable='false' type='flipboard' w='98426' x='787' y='14730'>
              <flipboard active-id='2' nav-type='caption' show-nav-arrows='true'>
                <story-points>""".format(story_sizing_string = story_sizing_string))

    for i in range(len(image_list)):
      f.write("""
                 <story-point captured-sheet='Dashboard {slide_num}' id='{slide_num}' />""".format(slide_num = i + 1))

    # Story trailer
    f.write("""
                </story-points>
              </flipboard>
            </zone>
          </zone>
          <zone-style>
            <format attr='border-color' value='#000000' />
            <format attr='border-style' value='none' />
            <format attr='border-width' value='0' />
            <format attr='margin' value='8' />
          </zone-style>
        </zone>
      </zones>
      <simple-id uuid='{503D6677-4C88-47BE-9B70-D9B6504FB60B}' />
    </dashboard>""")

    f.write("""
  </dashboards>
  <windows>""")
    
    ### Create a unique id for each window created (per dashboard and per slide)
    for i in range(len(image_list)):
      write_story_point(f, i + 1)

    f.write("""
    <window class='dashboard' maximized='true' name='Story 1'>
      <viewpoints />
      <active id='-1' />
      <simple-id uuid='{C8E3C7C3-8B64-490D-8564-A7EA63E551AE}' />
    </window>""")

    f.write("""
  </windows>
</workbook>""")

  if args.open:
    subprocess.call(['open', tableau_file_path])

if __name__ == "__main__":

  parser = argparse.ArgumentParser(description = "Creates a Tableau workbook with a story containing one storypoint for each image in a folder")

  parser.add_argument("images_folder_path", help = "Absolute pathname of the folder containing images to include in the Tableau workbook")
  parser.add_argument("tableau_path_name", help = "Pathname of the Tableau workbook to create, .twb extension optional")
  parser.add_argument("-f", "--fixed", metavar = ('HEIGHT', 'WIDTH'), help = "Use a fixed size for dashboards and storypoints (the default is an automatic/responsive size). Requires a height and width size in pixels.", type = int, nargs = 2, action = "append")
  parser.add_argument("-o", "--open", help = "Open the generated workbook after creating it.", action = "store_true")
  parser.add_argument("-r", "--replace", help = "Replaces a Tableau workbook if one already exists with the same name.", action = "store_true")

  create_story_workbook(parser.parse_args())

