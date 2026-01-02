#!/bin/python
import csv
import glob
import os
import shutil
import subprocess
from collections import namedtuple
from os import path

import click


class ScriptError(BaseException):
  def __init__(self, title: str, details: str, error_code: int):
    self._title = title
    self._details = details
    self._error_code = error_code\
    
  @property
  def title(self):
    return self._title

  @property
  def details(self):
    return self._details

  @property
  def error_code(self):
    return self._error_code

def probe_kicad_path() -> str: 
  KicadProbeList = ['/Applications/KiCad/KiCad.app/Contents/MacOS/kicad-cli']
  return next(iter([test_path for test_path in KicadProbeList if path.isfile(test_path)]), 'kicad-cli')

def run_command(command_args: list[str]):
  created_process = subprocess.run(command_args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
  if created_process.returncode != 0:
    raise ScriptError(f'Error executing command {" ".join(command_args)}', created_process.stdout.decode(), 2)

def script_step(title: str):
  def decorator(fn):
    fn.__step_title = title
    return fn
  
  return decorator

ScriptExecutionContext = namedtuple('ScriptExecutionContext', [
  'kicad_path', 
  'output', 
  'pcb', 
  'schema', 
  'project_base',
  'pos_fixups'])

@script_step('Generate gerbers')
def generate_gerbers(ec: ScriptExecutionContext):
  run_command([ec.kicad_path, 
         'pcb', 'export', 'gerbers', 
         '-l', 'F.Cu,F.Paste,F.Silkscreen,F.Mask,B.Cu,B.Paste,B.Silkscreen,B.Mask,Edge.Cuts', 
         '--no-x2', 
         '--subtract-soldermask', 
         '-o', ec.output, 
         ec.pcb])


@script_step('Generate drill')
def generate_drill(ec: ScriptExecutionContext):
      run_command([ec.kicad_path,
         'pcb', 'export', 'drill', 
         '--map-format', 'gerberx2', 
         '-o', f'{ec.output}/', 
         ec.pcb])

@script_step('Archive PCB fabrication outputs')
def archive_pcb(ec: ScriptExecutionContext):
      run_command(['zip',
                   '-o', f'{ec.output}/{ec.project_base}-gerbers.zip'] + 
                  glob.glob(f'{ec.output}/*.g*') + 
                  glob.glob(f'{ec.output}/*.drl'))

@script_step('Generate BOM')
def generate_bom(ec: ScriptExecutionContext):
    run_command(    
        [ec.kicad_path, 
         'sch', 'export', 'bom', 
          '--fields', 'Value,Reference,Footprint,LCSC', 
          '--group-by', 'Value', 
          '--labels', 'Comment,Designator,Footprint,LCSC Part Number', 
          '--ref-range-delimiter', '', 
           '-o', f'{ec.output}/{ec.project_base}-bom.csv', 
           ec.schema])

@script_step('Load pick-and-place fixup mappings')
def load_pos_fixups(ec: ScriptExecutionContext):
  pos_fixups_filename = f'{ec.output}/pos-fixups.csv' 
  try:
    run_command(
      [ec.kicad_path, 
      'sch', 'export', 'bom', 
        '--fields', 'Reference,PosRotAdjust',
        '-o', pos_fixups_filename, ec.schema])    
    with open(pos_fixups_filename) as pos_fixup_mappings_file:
      for fixup in csv.DictReader(pos_fixup_mappings_file):
        ec.pos_fixups[fixup['Reference']] = fixup

  finally:
    os.remove(pos_fixups_filename)

@script_step('Generate and fixup pick-and-place')
def generate_pos(ec: ScriptExecutionContext):
  pre_fixup_pos_filename = f'{ec.output}/pre-fixup-pos.pos'
  try:
    run_command(      
      [ec.kicad_path, 
      'pcb', 'export', 'pos', 
      '--format', 'csv',
        '--units', 'mm',
        '--side', 'front',
          '-o', pre_fixup_pos_filename,
          ec.pcb])
    with open(pre_fixup_pos_filename) as pre_fixup_pos_file:
      pre_fixup_pos_reader = csv.reader(pre_fixup_pos_file)
      next(pre_fixup_pos_reader)

      with open(f'{ec.output}/{ec.project_base}.pos', "w") as pos_file:
        pos_file.write(",".join(["Designator", "Val", "Package", "Mid X", "Mid Y", "Rotation", "Layer"]))            
        pos_file.write("\n")

        for pos_line in pre_fixup_pos_reader:
          if fixup := ec.pos_fixups[pos_line[0]]:
            pos_line[5] = fixup['PosRotAdjust'] or pos_line[5]

          pos_line[0] = f'"{pos_line[0]}"'
          pos_line[1] = f'"{pos_line[1]}"'
          pos_line[2] = f'"{pos_line[2]}"'
          
          pos_file.write(",".join(pos_line))
          pos_file.write("\n")

  finally:
    os.remove(pre_fixup_pos_filename)

@click.command()
@click.option('--project', '-p', required=True, type=click.Path(exists=True), help='KiCad project folder to generate fabraication output for')
@click.option('--schema', '-s', help='KiCad schema file, within the project, for fabrication')
@click.option('--pcb', '-c', help='KiCad PCB file, within the project, to generate fabraication for')
@click.option('--output', '-o', default='fab', help='Generation output fodler')
@click.option('--force', '-f', is_flag=True, help="Delete and recreate output folder if it doesn't exist")
@click.option('--kicad-path', '-k', help='Path to KiCad')
def generate_fab(project: str, schema: str, pcb: str, output: str, force: bool, kicad_path: str):
  try:
    if not kicad_path:
      kicad_path = probe_kicad_path()

    if not schema:
      schema = f'{path.basename(project)}.kicad_sch'

    if not pcb:
      pcb = f'{path.basename(project)}.kicad_pcb'
    
    schema = path.join(project, schema)
    pcb = path.join(project, pcb)

    if path.isdir(output):
      if force:
        shutil.rmtree(output)
      else:
        raise ScriptError(
          f'"{output}" already exists.', 
          'Use --fore to remove and re-create, or --output to specify a different path', 
          1)
    
    project_base = path.basename(project)

    os.makedirs(output)

    click.echo(f'Using {kicad_path} to generate fabraction outputs for PCB {pcb} and schema {schema}')

    ec = ScriptExecutionContext(kicad_path, output, pcb, schema, project_base, dict())

    with click.progressbar([
      generate_gerbers,
      generate_drill,
      archive_pcb,
      generate_bom,
      load_pos_fixups,
      generate_pos],
      item_show_func=lambda x: x.__step_title if x else None) as steps:
      for step in steps:
        step(ec)
      
    click.echo('Done.')
  except ScriptError as se:
    click.echo(click.style(se.title, fg='red'))
    click.echo(se.details)
    return se.error_code

if __name__ == '__main__':
  generate_fab()  
