import re
import sys
import requests
from urllib.parse import urlparse
import xlwt
import os
import zipfile
from xlwt import Workbook
import openpyxl
from openpyxl.styles import Border, Side

# Function to check if a given URL is valid
def is_valid_url(url):
    try:
        # Parse the URL using urlparse to extract its components
        result = urlparse(url)
        # Check if both scheme (protocol) and netloc (domain) are present
        return all([result.scheme, result.netloc])
    except ValueError:
        # If an exception occurs (invalid URL format), return False
        return False
    
# Function to get content from a given source (URL or file)
def get_content(input_source):
    if is_valid_url(input_source):
        # If the input source is a valid URL
        # Send an HTTP GET request to the URL
        response = requests.get(input_source)
        # Get the content (HTML/text) from the response
        content = response.text
    else:
        # If the input source is not a valid URL (assumed to be a file path)
        # Open the file in read mode and read its content
        with open(input_source, 'r') as f:
            content = f.read()
    return content

def extract_ips_domains_urls_files(content):
    ip_regex = r'\b(?:\d{1,3}\[?\.\]?){3}\d{1,3}\b'
    domain_regex = r'\b((?:[\w-]+\.)+(?:aaa|aarp|abarth|abb|abbott|abbvie|abc|abudhabi|ac|academy|accenture|accountant|accountants|aco|active|actor|ad|ads|adult|ae|aeg|aero|aetna|af|afl|africa|ag|agakhan|agency|ai|aig|aigo|airbus|airforce|airtel|akdn|al|alfaromeo|alibaba|alipay|allfinanz|allstate|ally|alsace|alsace|alstom|am|amazon|americanexpress|amex|amica|amsterdam|an|analytics|android|anz|ao|aol|apartments|app|apple|aq|aquarelle|ar|arab|aramco|archi|army|art|arte|as|asia|associates|at|attorney|au|auction|audi|audible|audio|auspost|author|auto|autos|aw|aws|ax|axa|az|azure|ba|baby|baidu|bananarepublic|band|bank|bar|barcelona|barclaycard|barclays|barefoot|bargains|baseball|basketball|bauhaus|bayern|bb|bbc|bbs|bbt|bbva|bcg|bcn|bd|be|beauty|beer|bentley|berlin|best|bestbuy|bet|bf|bg|bh|bharti|bi|bible|bid|bike|bing|bingo|bio|bit|biz|bj|bl|black|blackfriday|blanco|blockbuster|blog|bloomberg|blue|bm|bms|bmw|bn|bnl|bnpparibas|bnpparibas4|bo|boehringer|bond|boo|book|booking|boots|bosch|bostik|boston|bot|boutique|box|bq|br|bradesco|bridgestone|broadway|broker|brother|brussels|brussels|bs|bt|bu|bugatti|build|builders|business|buy|buzz|bv|bw|by|bz|bzh|bzh|ca|cab|cafe|cal|call|calvinklein|cam|camera|camp|cancerresearch|canon|capetown|capital|capitalone|car|caravan|cards|care|career|careers|cars|cartier|case|cash|casino|cat|cat|catering|catholic|cba|cbn|cbre|cbs|cc|cd|center|ceo|cern|cf|cfa|cfd|cg|ch|chanel|channel|charity|chase|chat|cheap|chintai|christmas|chrome|chrysler|church|ci|cipriani|circle|cisco|citadel|citi|citic|city|ck|cl|claims|cleaning|click|clinic|clothing|cloud|club|clubmed|cm|cn|co|coach|codes|coffee|college|cologne|com|comac|comad|comae|comaf|comag|comai|comal|comalsace|comam|coman|comao|comaq|comar|comas|comat|comau|comaw|comax|comaz|comba|combb|combd|combe|combf|combg|combh|combi|combj|combl|combm|combn|combo|combq|combr|combrussels|combs|combt|combu|combv|combw|comby|combz|combzh|comca|comcat|comcc|comcd|comcf|comcg|comch|comci|comck|comcl|comcm|comcn|comco|comcorsica|comcr|comcs|comcu|comcv|comcw|comcx|comcy|comcz|comdd|comde|comdj|comdk|comdm|comdo|comdz|comec|comee|comeg|comeh|comer|comes|comet|comeu|comfi|comfj|comfk|comfm|comfo|comfr|comga|comgb|comgd|comge|comgf|comgg|comgh|comgi|comgl|comgm|comgn|comgp|comgq|comgr|comgs|comgt|comgu|comgw|comgy|comhk|comhm|comhn|comhr|comht|comhu|comid|comie|comil|comim|comin|comio|comiq|comir|comis|comit|comje|comjm|comjo|comjp|comke|comkg|comkh|comki|comkm|comkn|comkp|comkr|comkrd|comkw|comky|comkz|comla|comlb|comlc|comli|comlk|comlr|comls|comlt|comlu|comlv|comly|comma|commc|commd|comme|commf|commg|commh|commk|comml|commm|commn|commo|commp|commq|commr|comms|commt|commu|commv|commw|commx|commy|commz|comna|comnc|comne|comnf|comng|comni|comnl|comno|comnp|comnr|comnu|comnz|comom|compa|compe|compf|compg|comph|compk|compl|compm|compn|compr|comps|compt|compw|compy|comqa|comqc|comre|comro|comrs|comru|comrw|comsa|comsb|comsc|comsd|comse|comsg|comsh|comsi|comsj|comsk|comsl|comsm|comsn|comso|comsr|comss|comst|comsu|comsv|comsx|comsy|comsz|comtc|comtd|comtf|comtg|comth|comtj|comtk|comtl|comtm|comtn|comto|comtp|comtr|comtt|comtv|comtw|comtz|comua|comug|comuk|comum|comus|comuy|comuz|comva|comvc|comve|comvg|comvi|comvlaanderen|comvn|comvu|comwf|comws|comye|comyt|comyu|comza|comzm|comzr|comzw|comcast|commbank|community|company|compare|computer|condos|construction|consulting|contact|contractors|cooking|cool|coop|corsica|corsica|country|coupon|coupons|courses|cpa|cr|credit|creditcard|creditunion|cricket|crown|crs|cruise|cruises|crypto,|cs|csc|cu|cuisinella|cv|cw|cx|cy|cymru|cyou|cz|dabur|dad|dance|data|date|dating|datsun|day|dd|de|deal|dealer|deals|degree|delivery|dell|deloitte|delta|democrat|dental|dentist|design|dev|dhl|diamonds|diet|digital|direct|directory|discount|discover|dish|diy|dj|dk|dm|dnp|do|docs|doctor|dodge|dog|doha|domains|dot|download|drive|dubai|dunlop|dupont|durban|dvag|dz|earth|eat|ec|eco|edeka|edu|eduua|education|ee|eg|eh|email|emerck|energy|engineer|engineering|enterprises|entertainment|epson|equipment|er|ericsson|erni|es|esq|estate|esurance|et|eth|etisalat|eu|eurovision|eus|events|everbank|example|exchange|expert|exposed|express|extraspace|fage|fail|fairwinds|faith|family|fan|fans|farm|farmers|fashion|fast|fedex|feedback|ferrari|ferrero|fi|fiat|fidelity|film|final|finance|financial|fire|firestone|firmdale|fish|fishing|fit|fitness|fj|fk|flickr|flights|flir|florist|flowers|flsmidth|fly|fm|fo|foo|food|foodnetwork|football|ford|forex|forsale|forum|foundation|fox|fr|free|fresenius|frl|frogans|frontdoor|frontier|fujitsu|fujixerox|fun|fund|fur|furniture|fyi|ga|gal|gallery|gallo|gallup|game|games|gap|garden|gay|gb|gbiz|gd|gdn|ge|gea|gent|genting|gf|gg|gh|gi|gift|gifts|gives|giving|gl|glass|gle|global|globo|gm|gmail|gmo|gmx|gn|godaddy|gold|goldpoint|golf|goodyear|goog|google|gop|gov|govng|govua|gp|gq|gr|grainger|graphics|green|gripe|grocery|group|gs|gt|gu|guardian|gucci|guide|guitars|guru|gw|gy|hair|hamburg|hangout|hbo|hdfc|hdfcbank|health|healthcare|help|helsinki|here|hermes|hiphop|hisamitsu|hitachi|hiv|hk|hkt|hm|hn|hockey|holdings|holiday|homegoods|homes|homesense|honda|honeywell|horse|hospital|host|hosting|hot|hotels|hotmail|house|how|hr|hsbc|ht|hu|hughes|hyatt|hyundai|ibm|ice|icu|id|ie|ieee|ifm|ikano|il|im|imdb|in|inc|industries|infiniti|info|ing|ink|institute|institute[45]|insurance|insure|int|intel|international|intuit|invalid|investments|io|ipiranga|iq|ir|irish|is|iselect|ist|istanbul|it|itau|itv|iveco|jaguar|java|jcb|jcp|je|jeep|jewelry|jm|jo|jobs|joburg|joy|jp|jpmorgan|juniper|kddi|ke|kerryhotels|kerrylogistics|kerryproperties|kfh|kg|kh|ki|kia|kim|kinder|kindle|kitchen|kiwi|km|kn|koeln|komatsu|kosher|kp|kpmg|kpn|kr|krd|krd|kred|ku|kuokgroup|kw|ky|kyoto|kz|la|lacaixa|ladbrokes|lamborghini|lancaster|lancia|lancome|land|landrover|lanxess|lasalle|lat|latrobe|law|lawyer|lb|lc|lds|lease|leclerc|legal|lego|lexus|lgbt|li|liaison|lib|lidl|life|lifeinsurance|lifestyle|lighting|like|lilly|limited|limo|lincoln|linde|link|lipsy|live|living|lixil|lk|loan|loans|local|localhost|locker|locus|lol|london|lotte|lotto|love|lpl|lplfinancial|lr|ls|lt|ltd|lu|lundbeck|lupin|luxury|lv|ly|ma|macys|madrid|maif|makeup|man|management|mango|map|market|marketing|markets|marriott|maserati|mattel|mba|mc|mckinsey|md|me|med|media|meet|melbourne|meme|memorial|men|menu|metlife|mf|mg|mh|miami|microsoft|mil|mini|mint|mit|mitsubishi|mk|ml|mlb|mm|mma|mn|mo|mobi|mobile|mobily|moe|mom|monash|money|monster|mormon|mortgage|moscow|moto|motorcycles|mov|movie|movistar|mp|mq|mr|ms|msd|mt|mtn|mtr|mu|museum|music|mutual|mv|mw|mx|my|mz|na|nadex|nagoya|name|nationwide|natura|navy|nba|nc|ne|nec|net|netua|netflix|network|neustar|new|newholland|news|nexus|nf|nfl|ng|ngo|nhk|ni|nico|nike|nikon|ninja|nissan|nissay|nl|no|nokia|northwesternmutual|norton|now|np|nr|nra|nrw|ntt|nu|nyc|nz|obi|observer|office|okinawa|om|omega|one|ong|onion|onl|online|ooo|open|oracle|orange|org|orgua|organic|origins|osaka|otsuka|ovh|pa|page|panasonic|paris|partners|parts|party|pay|pccw|pe|pet|pf|pfizer|pg|ph|pharmacy|philips|phone|photo|photography|photos|physio|piaget|pics|pictet|pictures|pid|pin|ping|pink|pioneer|pizza|pk|pl|place|play|playstation|plumbing|plus|pm|pn|pohl|poker|politie|porn|post|pr|praxi|press|prime|pro|prod|productions|prof|progressive|promo|properties|property|protection|pru|prudential|ps|pt|pub|pw|pwc|py|qa|qc|qpon|quebec|quest|qvc|racing|radio|re|read|realestate|realtor|realty|recipes|red|redstone|rehab|reit|reliance|rent|rentals|repair|report|republican|rest|restaurant|review|reviews|rexroth|rich|ricoh|rio|rip|rmit|ro|rocher|rocks|rodeo|rogers|room|rs|ru|rugby|ruhr|run|rw|rwe|ryukyu|sa|saarland|safe|safety|sakura|sale|salon|samsung|sandvik|sandvikcoromant|sanofi|sap|save|saxo|sb|sbi|sbs|sc|sca|scb|schaeffler|schmidt|scholarships|school|schwarz|science|scjohnson|scor|scot|sd|se|search|seat|secure|security|seek|select|sener|services|ses|seven|sew|sex|sexy|sfr|sg|sh|shangrila|sharp|shaw|shell|shiksha|shoes|shop|shopping|show|showtime|shriram|si|silk|sina|singles|site|sj|sk|ski|skin|sky|skype|sl|sling|sm|smart|smile|sn|sncf|so|soccer|social|softbank|software|sohu|solar|solutions|song|sony|sony12|spa|space|spiegel|sport|spot|spreadbetting|sr|srl|ss|st|stada|staples|star|starhub|statebank|statefarm|statoil|stc|stcgroup|stockholm|storage|store|stream|studio|study|style|su|sucks|supplies|supply|support|surf|surgery|suzuki|sv|swatch|swiftcover|swiss|sx|sy|sydney|symantec|systems|sz|taipei|talk|taobao|target|tatamotors|tatar|tattoo|tax|taxi|tc|td|tdk|team|tech|technology|tel|telecity|telefonica|temasek|tennis|test|teva|tf|tg|th|theater|theatre|tickets|tiffany|tips|tires|tirol|tj|tjx|tk|tl|tm|tn|to|today|tokyo|tools|top|toray|toshiba|total|tours|town|toyota|toys|tp|tr|trade|trading|training|travel|travelchannel|travelers|travelersinsurance|trust|tt|tt|tube|tui|tunes|tv|tvs|tw|tz|ua|ubs|uconnect|ug|uk|um|unicom|university|uno|uol|ups|us|uy|uz|va|vacations|vanguard|vc|ve|vegas|ventures|verisign|vet|vg|vi|video|vig|viking|villas|vin|vip|virgin|visa|vision|vista|vistaprint|vivo|vlaanderen|vlaanderen|vn|vodka|volkswagen|volvo|vote|voting|voyage|vu|wales|walmart|walter|wang|watch|watches|weather|weatherchannel|webcam|weber|website|wed|wedding|weir|wf|whoswho|wien|wiki|williamhill|win|windows|wine|winners|wme|wolterskluwer|woodside|work|works|world|wow|ws|wtc|wtf|xbox|xerox|xfinity|xn--3ds443g|xn--6frz82g|xn--fiq228c5hs|xn--q9jyb4c|xxx|xyz|yachts|yahoo|yamaxun|yandex|ye|yodobashi|yoga|yokohama|you|youtube|yt|yu|za|zappos|zara|zero|zippo|zm|zone|zr|zuerich|zw|みんな|中文网|在线|移动)\b)'
    url_regex = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
    file_regex = r'(?<=\W)([\w\-éèà^ôîç@#.$^~!@#%&+]+)\.(?:7z|a|aac|ace|alz|apk|appx|at3|arc|arj|b|ba|bin|bkf|blend|bz2|bmp|cab|c4|cals|xaml|cpt|sea|daa|deb|dmg|ddz|dn|dng|dpe|egg|egt|ecab|ezip|esd|ess|exe|flipchart|gbs|ggp|gsc|gho|ghs|gif|gz|html|ipg|jar|jpg|jpeg|lawrence|lbr|lqr|lzh|lz|lzo|lzma|lzx|lua|mbw|mhtml|midi|mpq|nl2pkg|nth|oar|osg|osk|osr|osz|pak|par|par2|paf|pea|png|webp|php|pyk|pk3|pk4|pxz|py|pyw|rar|rag|rags|rax|rbxl|rbxlx|rbxm|rbxmx|rpm|sb|sb2|sb3|sen|sitx|sis|sisx|skb|sq|srt|swm|szs|tar|gzip|targz|tb|tib|uha|uue|viv|vol|vsa|wax|wim|xap|xz|z|zoo|zip|zim|iso|nrg|img|adf|adz|dms|dsk|d64|sdi|mds|mdx|cdi|cue|cif|c2d|b6t|b5t|bwt|ffppkg|lemonapp|msi|vdhx|3dxml|3mf|acp|amf|aec|ar|art|asc|asm|bim|brep|c3d|c3p|ccc|ccm|ccs|cad|catdrawing|catpart|catproduct|catprocess|cgr|ckd|ckt|co|drw|dft|dgn|dgk|dmt|dxf|dwb|dwf|dwg|easm|edrw|emb|eprt|escpcb|escsch|esw|excellon|exp|f3d|fcstd|fm|fmz|g|gbr|glm|grb|gri|gro|iam|icd|idw|ifc|iges|cel|io|ipn|ipt|jt|mcd|mdg|model|ocd|pipe|pln|prt|psm|psmodel|pwi|pyt|skp|rlf|rvm|rvt|rfa|rxf|s12|scad|scdoc|sldasm|slddrw|sldprt|dotxsi|step|stl|std|tct|tcw|unv|vc6|vlm|vs|wrl|x_b|x_t|xe|zofzproj|brd|bsdl|cdl|cpf|def|dspf|edif|fsdb|gdsii|hex|lef|lib|ms12|oasis|openaccess|psf|psfxl|sdc|sdf|spef|spi|cir|srec|s19|sst2|stil|sv|s*p|tlf|upf|v|vcd|vhd|vhdl|wgl|4db|4dd|4dindy|4dindx|4dr|accdb|accde|adt|apr|box|chml|daf|dat|db|dbf|dta|eap|fdb|fp|fp3|fp5|fp7|frm|gdb|gtable|kexi|kexic|kexis|ldb|lirs|mda|mdb|adp|mde|mdf|myd|myi|ncf|nsf|ntf|nv2|odb|ora|pcontact|pdb|pdi|pdx|prc|sql|rec|rel|rin|sdb|sqlite|udl|wadata|waindx|wamodel|wajournal|wdb|wmdb|avro|parquet|orc|ai|ave|zave|cdr|chp|pub|sty|cap|vgr|dtp|gdraw|ildoc|indd|mcf|pdf|pmd|ppp|psd|qxd|sla|scd|xcf|1st|600|602|abw|acl|afp|ami|ans|aww|ccf|csv|cwk|dbk|dita|doc|docm|docx|dot|dotx|dwd|epub|ezw|fdx|ftm|ftx|gdoc|hwp|hwpml|log|lwp|mbp|md|me|mcw|mobi|nb|nbp|neis|nt|nq|odm|odoc|odt|osheet|ott|omm|pages|pap|per|pdr|pdax|quox|radix-64|rtf|rpt|sdw|se|stw|sxw|tex|info|troff|txt|uof|uoml|via|wpd|wps|wpt|wrd|wrf|wri|xhtml|xht|xml|xps|myo|myob|tax|ynab|ifx|ofx|qfx|qif|abf|afm|bdf|bmf|brfnt|fnt|fon|mgf|otf|pcf|pfa|pfb|pfm|fond|sfd|snf|tdf|tfm|ttf|ttc|ufo|woff|ifds|dem|e00|geotiff|gml|gpx|itn|mxd|ov2|shp|tab|dted|kml|3dt|aty|cag|fes|mgmf|mm|mmp|tpc|act|ase|gpl|pal|icc|icm|blp|bti|cd5|cit|cr2|clip|cpl|dds|dib|djvu|exif|grf|icns|ico|iff|ilbm|lbm|jng|jfif|jp2|jps|kra|max|miff|mng|msp|nitf|otb|pbm|pc1|pc2|pc3|pcx|pdn|pgm|pi1|pi2|pi3|pict|pct|pnm|pns|ppm|procreate|psb|pdd|psp|px|pxm|pxr|raw|rle|sct|sgi|rgb|int|bw|tga|targa|icb|vda|vst|pix|tiff|tif|ep|vtf|xbm|xpm|zif|3dv|awg|cgm|cmx|dp|drawio|e2d|eps|fs|odg|moviebyu|renderman|svg|3dmlw|x3d|sxd|tgax|v2d|vdoc|vsd|vsdx|vnd|wmf|emf|xar|3dmf|3dm|3ds|abc|ac|an8|aoi|b3d|block|bmd3|bdl4|brres|bfres|c4d|cal3d|ccp4|cfl|cob|core3d|ctm|dae|dff|dpm|dts|fact|fbx|glb|gltf|hec|iob|jas|jmesh|ldr|lwo|lws|lxf|lxo|m3d|ma|mb|mpd|md2|md3|md5|mesh|miobject|miparticle|mimodel|mm3d|mpo|mrc|nif|obj|off|ogex|ply|pov|r3d|rwx|sia|sib|smd|u3d|usd|usda|usdc|usdz|vim|vrml97|vue|vwx|wings|w3d|x|z3d|zbmx|alias|jnlp|lnk|appref-ms|nal|url|webloc|sym|desktop|harwell-boeing|mml|odf|sxm|8bf|aout|app|bac|bpl|bundle|class|coff|com|dcu|dll|dol|ear|elf|ipa|jeff|ko|list|mach-o|nlm|o|rll|s1es|so|vap|war|xbe|xcoff|xex|xpi|ocx|tlb|vbx|dvi|pld|pcl|ps|snp|xsl-fo|css|xslt|xsl|tpl|mnb|msg|org|pst|ost|sc2|gslides|key|keynote|odp|otp|pez|pot|pps|ppt|pptx|prz|sdd|shf|show|shw|slp|sspss|sti|sxi|thmx|watch|mpp|bib|enl|ris|fits|silo|spc|eas3|eossa|hitran|root|csdm|netcdf|hdr|hdf|h4|h5|sdxf|cdf|cgns|fmf|grib|bufr|pp|nasa-ames|cml|mol|sd|dx|jdx|smi|g6|s6|ab1|asn1|bam|bcf|bed|caf|cram|ddbj|embl|fasta|fastq|gcproj|genbank|gff|gtf|maf|ncbi|nexus|nexml|nwk|phd|sam|sbml|scf|sff|sra|stockholm|swiss-prot|vcf|dcm|nifti|nii|niigz|gii|brik|head|mgh|mgz|minc|mnc|acq|adicht|bci2000|bkr|cfwb|dicom|ecgml|edf|edf+|fef|gdf|hl7aecg|mfer|openxdf|scp-ecg|sigif|wfdb|xdf|hl7|xdt|cbf|ebf|cbfx|ebfx|adb|ads|ahk|applescript|as|au3|awk|bat|bas|cljs|cmd|coffee|c|cia|cpp|cs|ino|erb|go|hta|ibi|ici|ijs|ipynb|itcl|js|jsfl|kt|m|nuc|nud|nut|nqp|pde|php?|pl|pm|ps1|ps1xml|psc1|psd1|psm1|pyc|pyo|r|raku|rakumod|rakudoc|rakutest|rb|rdp|red|rs|scpt|scptd|sdl|sh|spwn|syjs|sypy|tcl|tns|ts|vbs|xpl|ebuild|omf|gxk|ssh|ppk|nsign|cer|crt|der|p7b|p7c|p12|pfx|pem|axx|eea|tc|kode|nsigne|bpw|kdb|kdbx|cfg|gms|irock|sac|seed|mseed|segy|win|win32|8svx|16svx|aiff|aif|aifc|au|aup3|bwf|cdda|dsf|wav|cwav|ra|rm|flac|la|pac|ape|ofr|ofs|rka|shn|tak|thd|tta|wv|wma|bcwav|brstm|dtshd|dtsma|ast|aw|ac3|amr|mp1|mp2|mp3|spx|gsm|mpc|vqf|ots|swa|vox|voc|smp|ogg|mod|mt2|s3m|xm|it|mid|btm|darms|etf|gp|kern|ly|mei|mus|musx|mxl|mscx|mscz|smdl|niff|ptb|asf|cust|gym|jam|rmj|sid|txm|vgm|ym|pvd|aimppl|asx|ram|xspf|zpl|m3u|pls|als|alc|alp|atmos|audio|metadata|aup|band|cau|cpr|cwp|drm|dmkit|ens|flm|flp|grir|logic|mmr|mx6hs|npr|omfi|ptx|ptf|pts|rpp|rpp-bak|reapeaks|ses|sfk|sfl|sng|stf|snd|syn|ust|vcls|vpr|vsq|vsqx|🗿|dvr-ms|wtv|ada|2ada|1ada|s|bb|bmx|clj|cls|cbl|cc|cxx|cbp|csproj|d|dba|dbpro123|e|efs|el|for|ftn|f|f77|f90|frx|fth|ged|gm6|gmd|gmk|h|hpp|hxx|hs|i|inc|java|l|lgt|lisp|m4|ml|msqr|n|p|pas|php3|php4|php5|phps|phtml|piv|pli|pl1|prg|pro|pol|reds|resx|rc|rc2|rkt|rktl|scala|sci|sce|scm|sd7|skc|skd|skf|skg|ski|skk|skm|sko|skq|sks|skt|skz|sln|spin|stk|swg|vb|vbg|vbp|vip|vbproj|vcproj|vdproj|xq|y|ab2|ab3|aws|bcsv|clf|cell|gsheet|numbers|gnumeric|lcw|ods|qpw|slk|stc|sxc|vc|wk1|wk3|wk4|wks|wq1|xlk|xls|xlsb|xlsm|xlsx|xlr|xlt|xltm|xlw|tsv|dif|aaf|3gp|avchd|avi|bik|braw|cam|collab|flv|mpeg-1|mpeg-2|noa|fla|flr|sol|str|m4v|mkv|wrap|mov|mpeg|mpg|mpe|thp|mpeg-4|mxf|roq|nsv|svi|smk|swf|wmv|yuv|webm|drp|fcp|mswmm|ppj|prproj|imovieproj|veg|veg-bak|suf|wlmp|kdenlive|vpj|motn|imoviemobile|pds|vproj|mcaddon|mcfunction|mcmeta|mcpack|mcr|mctemplate|mcworld|nbs|gbx|replaygbx|challengegbx|mapgbx|systemconfiggbx|trackmaniavehiclegbx|vehicletuningsgbx|solidgbx|itemgbx|blockgbx|texturegbx|materialgbx|tmedclassicgbx|ghostgbx|controlstylegbx|scoresgbx|profilegbx|loc|scripttxt|deh|dsg|lmp|wad|bsp|map|mdl|pk2|fontdat|sav|u|uax|umx|unr|upk|usx|ut2|ut3|utx|uxx|dmo|grp|itm|sqf|sqm|pbo|lip|vmf|vmx|hl2|vpk|vmt|cgb|bol|dbpf|diva|esm|esp|hambu|he0|he2|he4|gcf|love|mca|nbt|oec|osb|osc|osf2|osu|osz2|p3d|plagueinc|pod|rct|rep|simcity|sc4lot|sc4model|smzip|solitairetheme8|usld|vvvvvv|cps|stm|pkg|chr|z5|scworld|scskin|scbtex|prison|escape|wbfs|gba|pss|a26|a52|a78|lnx|jag|j64|wdf|gcm|min|nds|dsi|gb|gbc|sgm|n64|v64|z64|u64|usa|jap|eur|pj|nes|fds|jst|fc#|gg|sms|sg|32x|smc|sfc|fig|srm|zst|zs1|zs2|zs3|zs4|zs5|zs6|zs7|zs8|zs9|z10|z11|z12|z13|z14|z15|z16|z17|z18|z19|z20|z21|z22|z23|z24|z25|z26|z27|z28|z29|z30|z31|z32|z33|z34|z35|z36|z37|z38|z39|z40|z41|z42|z43|z44|z45|z46|z47|z48|z49|z50|z51|z52|z53|z54|z55|z56|z57|z58|z59|z60|z61|z62|z63|z65|z66|z67|z68|z69|z70|z71|z72|z73|z74|z75|z76|z77|z78|z79|z80|z81|z82|z83|z84|z85|z86|z87|z88|z89|z90|z91|z92|z93|z94|z95|z96|z97|z98|z99|frz|pce|npc|ngp|ngc|vec|ws|wsc|tzx|tap|sna|t64|vfd|vud|vmc|vsv|vmdk|nvram|vmem|vmsd|vmsn|vmss|vmtm|vmxf|vdi|vbox-extpack|hdd|pvs|cow|qcow|qcow2|qed|dtd|htm|mht|maff|asp|aspx|bml|cfm|cgi|ihtml|jsp|las|lasso|lassoapp|shtml|atom|eml|jsonld|kprx|metalink|met|rss|markdown|axd|cex|col|credx|ddb|ddi|dupx|ftmb|ga3|hlp|igc|inf|kmc|kcl|ktr|lsm|narc|oer|pa|pif|por|rise|scr|topc|xlf|xmc|zed|zone|fx|miframes|milanguage|midata|bca|ani|cur|smes|ini|json|yaml|restructuredtext|asciidoc|yni|bak|bk|szh|cnf|conf|diff|!ut|crdownload|opdownload|part|partial|temp|tmp)(?=\W)'
    hashes = r'\b[A-Fa-f0-9]{32,}\b|\b[A-Fa-f0-9]{40,}\b|\b[A-Fa-f0-9]{64,}\b'
    bracketed_dot = r'\[\.\]'

    # Replace bracketed dots with normal dots (if provided in a report)
    content = re.sub(bracketed_dot, '.', content)

    domain_names = re.findall(domain_regex, content, re.IGNORECASE)
    domain_names = list(set(domain_names))

    for i in domain_names:
        if (i=='alert.aecert.ae'):
            domain_names.remove(i)

    for a in domain_names:
        if (a == 'emiratessteel.com'):
            domain_names.remove(a)
    for b in domain_names:
        if (b == 'aecert.ae'):
            domain_names.remove(b)


    
    # Extract file names while excluding any matches that also match with the extracted domain names
    file_names = []
    for match in re.finditer(file_regex, content):
        matched_file = match.group(0)
        if not any(domain for domain in domain_names if matched_file in domain):
            file_names.append(matched_file)

    # Extract other patterns
    ip_addresses = re.findall(ip_regex, content)
    urls = re.findall(url_regex, content)
    hashes = re.findall(hashes,content)
     
    # Deduplicate results
    ip_addresses = list(set(ip_addresses))
    urls = list(set(urls))

    for k in urls:
        if (k == 'https://alert.aecert.ae'):
            urls.remove(k)
    file_names = list(set(file_names))
    for j in file_names:
        if (j == 'aecert@emiratessteel.com'):
            file_names.remove(j)

    hashes = list(set(hashes))

    return ip_addresses, domain_names, urls, hashes, file_names



def create_excel():
    wb = Workbook()
    #Create sheet for IOCs
    sheet1 = wb.add_sheet('IPs')
    style = xlwt.easyxf('font: bold 1')
    sheet1.write(1, 1, 'IPs', style)
    i=1
    j=0

    # border_style = Border(
    # left=Side(style='thin'),
    # right=Side(style='thin'),
    # top=Side(style='thin'),
    # bottom=Side(style='thin'))

    # sheet1.border = border_style

    #style2=xlwt.easyxf('alignment: wrap True')
    for j in range(len(ips)):
        sheet1.write(i,2,ips[j])
        i=i+1

    #Create sheet for Hashes
    sheet2=wb.add_sheet('Hashes')
    style=xlwt.easyxf('font:bold 1')
    sheet2.write (1,1,'Hashes',style)
    a=1
    b=0
    for b in range(len(hashes)):
        sheet2.write(a,2,hashes[b])
        a=a+1

    #Create sheet for Domains
    sheet3=wb.add_sheet('Domains')
    style=xlwt.easyxf('font:bold 1')
    sheet3.write (1,1,'Domains',style)
    d=1
    e=0
    for e in range(len(domains)):
        sheet3.write(d,2,domains[e])
        d=d+1

    wb.save("Try.xls")

    # Function to apply a thin border to cells within specified rows and columns of a worksheet
def apply_border(worksheet, rows, cols):
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))

    for row in worksheet.iter_rows(min_row=rows[0], max_row=rows[1], min_col=cols[0], max_col=cols[1]):
        for cell in row:
            cell.border = thin_border

# Function to apply center and middle alignment to cells within specified rows and columns of a worksheet
def apply_alignment(worksheet, rows, cols):
    alignment = Alignment(horizontal='center', vertical='center')
    
    for row in worksheet.iter_rows(min_row=rows[0], max_row=rows[1], min_col=cols[0], max_col=cols[1]):
        for cell in row:
            cell.alignment = alignment


def create_zip(excel_path):
    
    output_zip_path = os.path.splitext(excel_path)[0] + ".zip"
    
    with zipfile.ZipFile(output_zip_path, "w") as zipf:
        zipf.write(excel_path)
    
    return output_zip_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 extract_iocs.py <path_to_file_or_url>")
        sys.exit(1)

    input_source = sys.argv[1]
    content = get_content(input_source)
    ips, domains, urls, hashes, files = extract_ips_domains_urls_files(content)

    print("\nIP Addresses:")
    print(type(ips))
    for ip in ips:
        print(ip)

    print("\nDomain Names:")
    for domain in domains:
        print(domain)

    print("\nURLs:")
    for url in urls:
        print(url)

    print("\nFile Names:")
    for file in files:
        print(file)
    
    print("\nHashes:")
    for hash in hashes:
        print(hash)
create_excel()
s_path=".\Try.xls"
x=create_zip(s_path)

print(f'Zip has been created as {x}')