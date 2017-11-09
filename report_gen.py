from docx import Document
from docx.shared import Cm
import datetime
import argparse
try:
    from halo import Halo
except:
    pass


parser = argparse.ArgumentParser(description='Generate pretty Strucutre Reports.')
parser.add_argument('-u','--user', help='name of the synthetic chemist' )
parser.add_argument('-i','--input', help='name of input CIF-file' )
parser.add_argument('-s','--special', help='special details for the refinement' )
parser.add_argument('-d','--docx', help='path to the template docx file' )

now = datetime.datetime.now()

keys_report = ['ID', 'USER', 'DATE', 'SUM', 'SUM1', 'MOI', 'WEIGHT', 'CRYSSYSTEM', 'SPGR', 'CELLA', 'CELLB',
    'CELLC', 'ALPHA', 'BETA', 'GAMMA', 'VOLUME', 'ZERR', 'DENSITY', 'MU', 'COMPLETENESS', 'THETAMIN',
    'THETAMAX', 'COLLRFL', 'UNIQUE', 'UNIQUE2S', 'ABSCORR', 'TMAX', 'TMIN', 'DATA', 'PARAM', 'RESTR', 'GOOF',
    'R1ALL', 'R12SIG', 'WR2ALL', 'WR2SIG', 'HIGHPEAK', 'DEEPHOLE', 'SIZEMAX', 'SIZEMIN', 'SIZEMID']

keys_cif = {'_cell_length_a':'CELLA', '_cell_length_b':'CELLB', '_cell_length_c':'CELLC', '_cell_angle_alpha':'ALPHA',
    '_cell_angle_beta':'BETA', '_cell_angle_gamma':'GAMMA', '_cell_volume':'VOLUME', '_cell_formula_units_Z':'ZERR',
    '_exptl_crystal_density_diffrn':'DENSITY', '_exptl_crystal_size_max':'SIZEMAX', '_exptl_crystal_size_mid':'SIZEMID',
    '_exptl_crystal_size_min':'SIZEMIN', '_exptl_absorpt_coefficient_mu':'MU', '_exptl_absorpt_correction_type':'ABSCORR',
    '_exptl_absorpt_correction_T_min':'TMIN','_exptl_absorpt_correction_T_max':'TMAX', '_diffrn_reflns_theta_min':'THETAMIN',
    '_diffrn_reflns_theta_max':'THETAMAX', '_diffrn_reflns_number':'COLLRFL', '_reflns_number_total':'UNIQUE',
    '_reflns_number_gt':'UNIQUE2S', '_diffrn_measured_fraction_theta_max':'COMPLETENESS', '_chemical_formula_sum':'SUM',
    '_chemical_formula_moiety':'MOI','_space_group_crystal_system':'CRYSSYSTEM','_chemical_formula_weight':'WEIGHT',
    '_refine_ls_number_reflns':'DATA','_refine_ls_number_parameters':'PARAM', '_refine_ls_number_restraints':'RESTR',
    '_refine_ls_restrained_S_all':'GOOF','_refine_ls_R_factor_all':'R1ALL', '_refine_ls_R_factor_gt':'R12SIG',
    '_refine_ls_wR_factor_ref':'WR2ALL', '_refine_ls_wR_factor_gt': 'WR2SIG',
    '_refine_diff_density_max':'HIGHPEAK', '_refine_diff_density_min':'DEEPHOLE'}

indent = ['ALPHA', 'BETA', 'GAMMA', 'R1ALL', 'R12SIG', 'WR2ALL', 'WR2SIG']
rValues = []

def cifReader(cifFile):
    file = open(cifFile,'r').readlines()
    outDict = {}
    outDict['DATE'] = now.strftime("%d.%m.%Y")
    outDict['IDCODE'] = cifFile.split('_')[0]
    for line in file:
        try:
            tempKey = line.split()[0]
            if tempKey in keys_cif.keys():
                try:
                    value = ' '.join(line.split()[1:])
                except IndexError:
                    value = '-'
                outDict[keys_cif[tempKey]] = value
        except IndexError:
            pass
    return outDict

try:
    spinner = Halo(text='Please be patient. I am generating your report.',spinner='smiley', color='red')
    spinner.start()
except:
    pass

args = parser.parse_args()
template = Document(args.docx)
replacements = cifReader(args.input)
replacements['USER'] = args.user

if args.special:
    replacements['SPECIAL'] = args.special
else:
    replacements['SPECIAL'] = 'none'

replacements['SIZE'] = '{} x {} x {}'.format(replacements['SIZEMIN'],replacements['SIZEMID'],replacements['SIZEMAX'])

for line in template.paragraphs:
    for key in replacements.keys():
        print key
        tabStops = line.paragraph_format.tab_stops
        tabStops.add_tab_stop(Cm(1.0))
        tabStops.add_tab_stop(Cm(4.0))
        tabStops.add_tab_stop(Cm(8.0))
        tabStops.add_tab_stop(Cm(12.0))      
        inline = line.runs
        stop = False
        for i in inline:
            if key in i.text:
                newLine = i.text.replace(key, replacements[key])
                i.text = newLine
                stop = True
        if stop:
            break
                
template.save('out.docx')

try:
    spinner.stop()
    spinner.succeed('Report generated')
except:
    pass
