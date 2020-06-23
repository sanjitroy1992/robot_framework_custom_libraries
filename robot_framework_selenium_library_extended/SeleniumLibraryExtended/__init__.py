import inspect
from time import sleep
import wrapt
from selenium.common.exceptions import StaleElementReferenceException
from selenium import webdriver
from selenium.webdriver import ActionChains
from SeleniumLibrary import SeleniumLibrary
from SeleniumLibrary.base import keyword
from .utils import _SynchronizationKeywords
from . import webdrivermonkeypatches
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.chrome.options import Options
import os
import time
from platform import system
from PIL import Image
from io import BytesIO
if system() == 'Windows':
    from KeyboardLibrary.KeyboardLibrary import KeyboardLibrary
    import pyautogui
__version__ = "1.0"

@wrapt.decorator
def _rerun_on_defined_exceptions(wrapped, instance, args, kwargs):
    try:
        return wrapped(*args, **kwargs)
    except StaleElementReferenceException:
        sleep(0.5)
        return wrapped(*args, **kwargs)


class SeleniumLibraryExtended(
    _SynchronizationKeywords,
):
    __doc__ = SeleniumLibrary.__doc__

    def __init__(self, timeout=5.0, implicit_wait=0.0, run_on_failure='Capture Page Screenshot',
                 screenshot_root_directory=None, safe_mode=True):
        self.safe_mode = safe_mode
        SeleniumLibrary.__init__(self, timeout, implicit_wait, run_on_failure, screenshot_root_directory)
        if self.safe_mode:
            self._decorate_all_unbound_methods(_rerun_on_defined_exceptions, SeleniumLibraryExtended)

    def _decorate_all_unbound_methods(self, decorator, cls):
        for name, method in inspect.getmembers(cls, inspect.ismethod):
            if not name.startswith('_'):
                setattr(cls, name, decorator(method))

    @keyword
    def move_to_element_with_offset(self, locator, x_offset, y_offset):
        self._info("Simulating Mouse Over on element with offset '%s'" % locator)
        element = self._element_find(locator, True, False)
        if element is None:
            raise AssertionError("ERROR: Element %s not found." % locator)
        ActionChains(self._current_browser()) \
            .move_to_element_with_offset(element, x_offset, y_offset) \
            .perform()

    @keyword
    def draw_rectangle_over_element(self, locator, source_x_offset, source_y_offset, width, height):
        element = self._element_find(locator, True, False)
        target_x_offset = int(source_x_offset) + int(width)
        target_y_offset = int(source_y_offset) + int(height)
        if element is None:
            raise AssertionError("ERROR: Element %s not found." % locator)
        ActionChains(self._current_browser()) \
            .move_to_element_with_offset(element, source_x_offset, source_y_offset) \
            .click_and_hold() \
            .move_to_element_with_offset(element, target_x_offset, target_y_offset) \
            .release() \
            .perform()

    @keyword
    def create_firefox_webdriver_with_proxy(self, proxy_host, proxy_port):
        """
        Creates an instance of a Firefox WebDriver with a certain proxy.
        Examples:
        | Create Firefox Webdriver With Proxy | localhost | 8080 |
        """
        fp = webdriver.FirefoxProfile()
        fp.set_preference("network.proxy.type", 1)
        fp.set_preference("network.proxy.http", proxy_host)
        fp.set_preference("network.proxy.http_port", int(proxy_port))
        fp.update_preferences()
        driver = webdriver.Firefox(firefox_profile=fp)
        return self._cache.register(driver)

    @keyword
    def textfield_should_contain(self, locator, expected, message=''):
        """Verifies text field identified by `locator` contains text `expected`.

        `message` can be used to override default error message.

        Key attributes for text fields are `id` and `name`. See `introduction`
        for details about locating elements.
        """

        actual = (self._get_value(locator, 'text field')).encode('utf8')
        if not expected in actual.decode('utf8'):
            if not message:
                message = "Text field '%s' should have contained text '%s' " \
                          "but it contained '%s'" % (locator, expected, actual.decode('utf8'))
            raise AssertionError(message)
        self._info("Text field '%s' contains text '%s'." % (locator, expected))

    @keyword
    def select_from_group(self, locator, *params):
        """ Used to select one or more option from list by text.

         |Arguments|

         'locator' = List elements locator

         'param' = Option text.

         |Example|

         1. To select one option from list

         | ***TestCases*** |
         | Select From Group | ${select_measures} | misys |



         2. To select multiple options from list

         | ***Variables*** |
         | @{selectAllMeasuresParam} | AdjNotional | SS_i | MarketValue | SE_i | SD_i | Delta_i | Collateral |

         | ***TestCases*** |
         | Select From Group | ${select_measures} | @{selectAllMeasuresParam} |

        """
        GroupLocator = self.get_webelements(locator)
        NotFound = ""
        if GroupLocator:
            for param in params:
                isFound = 0
                for element in GroupLocator:
                    if ((element.text).strip() == param):
                        isFound = 1
                        element.click()
                        break
                if isFound == 0:
                    NotFound += param + ", "
            if NotFound != "":
                raise AssertionError("{} element not found in {}".format(NotFound, locator))
        else:
            self.failure_occurred()
            raise AssertionError("Group Locator : " + locator + " is not found")

    @keyword
    def activate_and_input_text(self, locator, text):
        """ Used to activate element specified by 'locator' and then enter 'text'.

         |Arguments|

         'locator' = Input field locator

         'text' = text to input into locator

         |Example|

         | Activate And Input Text | ${TextBoxLocator} | misys |"""

        caps = self._current_browser().capabilities
        browserName = caps['browserName']
        if browserName in ['firefox', 'ff']:
            element = self._element_find(locator, True, True)
            ActionChains(self._current_browser()).move_to_element(element).click().perform()
            self._current_browser().switch_to.active_element['value'].send_keys(text)
        else:
            self.double_click_element(locator)
            self._current_browser().switch_to.active_element.send_keys(text)

    @keyword
    def scroll_element_into_view(self, locator):
        """Scrolls an element from given ``locator`` into view.

        Arguments:
        - ``locator``: The locator to find requested element.

        Examples:
        | Scroll Element Into View | css=div.class |
        """
        if isinstance(locator, WebElement):
            element = locator
        else:
            self._info("Scrolling element '%s' into view." % locator)
            element = self._element_find(locator, True, True)
        script = 'arguments[0].scrollIntoView()'
        # pylint: disable=no-member
        self._current_browser().execute_script(script, element)
        return element

    @keyword
    def hover_and_click_element(self, locator):
        """ Used to hover to any element.If that element is not present then it will raise an Error

         Arguments : 'locator' is the element locator

         Example :
        | Hover And Click Element | ${elementLocator} |
        """

        MoreOptions = self._element_find(locator, True, True)
        if MoreOptions:
            ActionChains(self._current_browser()).move_to_element(MoreOptions).click().perform()
        else:
            raise AssertionError(locator + " Element Not Found")

    @keyword
    def set_firefox_profile_and_enable_download_directory(self, browserName, download_dir, headless_mode='No'):

        """Used to set firefox profile with the options to set the default download directory and to launch in headless mode.

        browsername : Browser name to be specified as firefox.
        download_dir : Directory of the path where the user wants to set the default download path.
                       Please see the below examples and give the directory path with two backslashes (\\\\)
        headless_mode : This enables user to launch the firefox in headless mode if set to yes,
                        By default headless mode is disabled. Please see the example 2.
        Example :
        1. To launch an url in UI mode with download directory set.
        | ${var1}  |  Set Firefox Profile And Enable Download Directory  |  firefox  |  D:\\\\Dummy_Download |
        | Open Browser  |  https://www.finastra.com  |  Firefox  |  ff_profile_dir=${var1} |

        2. To launch an url in headless mode with download directory set.
        | ${var1}  |  Set Firefox Profile And Enable Download Directory  |  firefox  |  D:\\\\Dummy_Download | headless_mode=Yes |
        | Open Browser  |  https://www.finastra.com  |  Firefox  |  ff_profile_dir=${var1} |
        | Set Window Size  |  1920  |  1080 |

        """
        if browserName.lower() in ['firefox', 'ff']:
            if headless_mode.lower() in ['yes']:
                os.environ['MOZ_HEADLESS'] = '1'
            fp = webdriver.FirefoxProfile()
            fp.set_preference("browser.download.folderList", 2)
            fp.set_preference("browser.download.manager.showWhenStarting", False)
            fp.set_preference("browser.download.dir", download_dir)
            fp.set_preference("browser.helperApps.neverAsk.saveToDisk",
                              "application/vnd.hzn-3d-crossword;video/3gpp;video/3gpp2;application/vnd.mseq;application/vnd.3m.post-it-notes;application/vnd.3gpp.pic-bw-large;application/vnd.3gpp.pic-bw-small;application/vnd.3gpp.pic-bw-var;application/vnd.3gp2.tcap;application/x-7z-compressed;application/x-abiword;application/x-ace-compressed;application/vnd.americandynamics.acc;application/vnd.acucobol;application/vnd.acucorp;audio/adpcm;application/x-authorware-bin;application/x-athorware-map;application/x-authorware-seg;application/vnd.adobe.air-application-installer-package+zip;application/x-shockwave-flash;application/vnd.adobe.fxp;application/pdf;application/vnd.cups-ppd;application/x-director;applicaion/vnd.adobe.xdp+xml;application/vnd.adobe.xfdf;audio/x-aac;application/vnd.ahead.space;application/vnd.airzip.filesecure.azf;application/vnd.airzip.filesecure.azs;application/vnd.amazon.ebook;application/vnd.amiga.ami;applicatin/andrew-inset;application/vnd.android.package-archive;application/vnd.anser-web-certificate-issue-initiation;application/vnd.anser-web-funds-transfer-initiation;application/vnd.antix.game-component;application/vnd.apple.installe+xml;application/applixware;application/vnd.hhe.lesson-player;application/vnd.aristanetworks.swi;text/x-asm;application/atomcat+xml;application/atomsvc+xml;application/atom+xml;application/pkix-attr-cert;audio/x-aiff;video/x-msvieo;application/vnd.audiograph;image/vnd.dxf;model/vnd.dwf;text/plain-bas;application/x-bcpio;application/octet-stream;image/bmp;application/x-bittorrent;application/vnd.rim.cod;application/vnd.blueice.multipass;application/vnd.bm;application/x-sh;image/prs.btif;application/vnd.businessobjects;application/x-bzip;application/x-bzip2;application/x-csh;text/x-c;application/vnd.chemdraw+xml;text/css;chemical/x-cdx;chemical/x-cml;chemical/x-csml;application/vn.contact.cmsg;application/vnd.claymore;application/vnd.clonk.c4group;image/vnd.dvb.subtitle;application/cdmi-capability;application/cdmi-container;application/cdmi-domain;application/cdmi-object;application/cdmi-queue;applicationvnd.cluetrust.cartomobile-config;application/vnd.cluetrust.cartomobile-config-pkg;image/x-cmu-raster;model/vnd.collada+xml;text/csv;application/mac-compactpro;application/vnd.wap.wmlc;image/cgm;x-conference/x-cooltalk;image/x-cmx;application/vnd.xara;application/vnd.cosmocaller;application/x-cpio;application/vnd.crick.clicker;application/vnd.crick.clicker.keyboard;application/vnd.crick.clicker.palette;application/vnd.crick.clicker.template;application/vn.crick.clicker.wordbank;application/vnd.criticaltools.wbs+xml;application/vnd.rig.cryptonote;chemical/x-cif;chemical/x-cmdf;application/cu-seeme;application/prs.cww;text/vnd.curl;text/vnd.curl.dcurl;text/vnd.curl.mcurl;text/vnd.crl.scurl;application/vnd.curl.car;application/vnd.curl.pcurl;application/vnd.yellowriver-custom-menu;application/dssc+der;application/dssc+xml;application/x-debian-package;audio/vnd.dece.audio;image/vnd.dece.graphic;video/vnd.dec.hd;video/vnd.dece.mobile;video/vnd.uvvu.mp4;video/vnd.dece.pd;video/vnd.dece.sd;video/vnd.dece.video;application/x-dvi;application/vnd.fdsn.seed;application/x-dtbook+xml;application/x-dtbresource+xml;application/vnd.dvb.ait;applcation/vnd.dvb.service;audio/vnd.digital-winds;image/vnd.djvu;application/xml-dtd;application/vnd.dolby.mlp;application/x-doom;application/vnd.dpgraph;audio/vnd.dra;application/vnd.dreamfactory;audio/vnd.dts;audio/vnd.dts.hd;imag/vnd.dwg;application/vnd.dynageo;application/ecmascript;application/vnd.ecowin.chart;image/vnd.fujixerox.edmics-mmr;image/vnd.fujixerox.edmics-rlc;application/exi;application/vnd.proteus.magazine;application/epub+zip;message/rfc82;application/vnd.enliven;application/vnd.is-xpr;image/vnd.xiff;application/vnd.xfdl;application/emma+xml;application/vnd.ezpix-album;application/vnd.ezpix-package;image/vnd.fst;video/vnd.fvt;image/vnd.fastbidsheet;application/vn.denovo.fcselayout-link;video/x-f4v;video/x-flv;image/vnd.fpx;image/vnd.net-fpx;text/vnd.fmi.flexstor;video/x-fli;application/vnd.fluxtime.clip;application/vnd.fdf;text/x-fortran;application/vnd.mif;application/vnd.framemaker;imae/x-freehand;application/vnd.fsc.weblaunch;application/vnd.frogans.fnc;application/vnd.frogans.ltf;application/vnd.fujixerox.ddd;application/vnd.fujixerox.docuworks;application/vnd.fujixerox.docuworks.binder;application/vnd.fujitu.oasys;application/vnd.fujitsu.oasys2;application/vnd.fujitsu.oasys3;application/vnd.fujitsu.oasysgp;application/vnd.fujitsu.oasysprs;application/x-futuresplash;application/vnd.fuzzysheet;image/g3fax;application/vnd.gmx;model/vn.gtw;application/vnd.genomatix.tuxedo;application/vnd.geogebra.file;application/vnd.geogebra.tool;model/vnd.gdl;application/vnd.geometry-explorer;application/vnd.geonext;application/vnd.geoplan;application/vnd.geospace;applicatio/x-font-ghostscript;application/x-font-bdf;application/x-gtar;application/x-texinfo;application/x-gnumeric;application/vnd.google-earth.kml+xml;application/vnd.google-earth.kmz;application/vnd.grafeq;image/gif;text/vnd.graphviz;aplication/vnd.groove-account;application/vnd.groove-help;application/vnd.groove-identity-message;application/vnd.groove-injector;application/vnd.groove-tool-message;application/vnd.groove-tool-template;application/vnd.groove-vcar;video/h261;video/h263;video/h264;application/vnd.hp-hpid;application/vnd.hp-hps;application/x-hdf;audio/vnd.rip;application/vnd.hbci;application/vnd.hp-jlyt;application/vnd.hp-pcl;application/vnd.hp-hpgl;application/vnd.yamaha.h-script;application/vnd.yamaha.hv-dic;application/vnd.yamaha.hv-voice;application/vnd.hydrostatix.sof-data;application/hyperstudio;application/vnd.hal+xml;text/html;application/vnd.ibm.rights-management;application/vnd.ibm.securecontainer;text/calendar;application/vnd.iccprofile;image/x-icon;application/vnd.igloader;image/ief;application/vnd.immervision-ivp;application/vnd.immervision-ivu;application/reginfo+xml;text/vnd.in3d.3dml;text/vnd.in3d.spot;mode/iges;application/vnd.intergeo;application/vnd.cinderella;application/vnd.intercon.formnet;application/vnd.isac.fcs;application/ipfix;application/pkix-cert;application/pkixcmp;application/pkix-crl;application/pkix-pkipath;applicaion/vnd.insors.igm;application/vnd.ipunplugged.rcprofile;application/vnd.irepository.package+xml;text/vnd.sun.j2me.app-descriptor;application/java-archive;application/java-vm;application/x-java-jnlp-file;application/java-serializd-object;text/x-java-source,java;application/javascript;application/json;application/vnd.joost.joda-archive;video/jpm;image/jpeg;video/jpeg;application/vnd.kahootz;application/vnd.chipnuts.karaoke-mmd;application/vnd.kde.karbon;aplication/vnd.kde.kchart;application/vnd.kde.kformula;application/vnd.kde.kivio;application/vnd.kde.kontour;application/vnd.kde.kpresenter;application/vnd.kde.kspread;application/vnd.kde.kword;application/vnd.kenameaapp;applicatin/vnd.kidspiration;application/vnd.kinar;application/vnd.kodak-descriptor;application/vnd.las.las+xml;application/x-latex;application/vnd.llamagraphics.life-balance.desktop;application/vnd.llamagraphics.life-balance.exchange+xml;application/vnd.jam;application/vnd.lotus-1-2-3;application/vnd.lotus-approach;application/vnd.lotus-freelance;application/vnd.lotus-notes;application/vnd.lotus-organizer;application/vnd.lotus-screencam;application/vnd.lotus-wordro;audio/vnd.lucent.voice;audio/x-mpegurl;video/x-m4v;application/mac-binhex40;application/vnd.macports.portpkg;application/vnd.osgeo.mapguide.package;application/marc;application/marcxml+xml;application/mxf;application/vnd.wolfrm.player;application/mathematica;application/mathml+xml;application/mbox;application/vnd.medcalcdata;application/mediaservercontrol+xml;application/vnd.mediastation.cdkey;application/vnd.mfer;application/vnd.mfmp;model/mesh;appliation/mads+xml;application/mets+xml;application/mods+xml;application/metalink4+xml;application/vnd.ms-powerpoint.template.macroenabled.12;application/vnd.ms-word.document.macroenabled.12;application/vnd.ms-word.template.macroenabed.12;application/vnd.mcd;application/vnd.micrografx.flo;application/vnd.micrografx.igx;application/vnd.eszigno3+xml;application/x-msaccess;video/x-ms-asf;application/x-msdownload;application/vnd.ms-artgalry;application/vnd.ms-ca-compressed;application/vnd.ms-ims;application/x-ms-application;application/x-msclip;image/vnd.ms-modi;application/vnd.ms-fontobject;application/vnd.ms-excel;application/vnd.ms-excel.addin.macroenabled.12;application/vnd.ms-excelsheet.binary.macroenabled.12;application/vnd.ms-excel.template.macroenabled.12;application/vnd.ms-excel.sheet.macroenabled.12;application/vnd.ms-htmlhelp;application/x-mscardfile;application/vnd.ms-lrm;application/x-msmediaview;aplication/x-msmoney;application/vnd.openxmlformats-officedocument.presentationml.presentation;application/vnd.openxmlformats-officedocument.presentationml.slide;application/vnd.openxmlformats-officedocument.presentationml.slideshw;application/vnd.openxmlformats-officedocument.presentationml.template;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.openxmlformats-officedocument.spreadsheetml.template;application/vnd.openxmformats-officedocument.wordprocessingml.document;application/vnd.openxmlformats-officedocument.wordprocessingml.template;application/x-msbinder;application/vnd.ms-officetheme;application/onenote;audio/vnd.ms-playready.media.pya;vdeo/vnd.ms-playready.media.pyv;application/vnd.ms-powerpoint;application/vnd.ms-powerpoint.addin.macroenabled.12;application/vnd.ms-powerpoint.slide.macroenabled.12;application/vnd.ms-powerpoint.presentation.macroenabled.12;appliation/vnd.ms-powerpoint.slideshow.macroenabled.12;application/vnd.ms-project;application/x-mspublisher;application/x-msschedule;application/x-silverlight-app;application/vnd.ms-pki.stl;application/vnd.ms-pki.seccat;application/vn.visio;video/x-ms-wm;audio/x-ms-wma;audio/x-ms-wax;video/x-ms-wmx;application/x-ms-wmd;application/vnd.ms-wpl;application/x-ms-wmz;video/x-ms-wmv;video/x-ms-wvx;application/x-msmetafile;application/x-msterminal;application/msword;application/x-mswrite;application/vnd.ms-works;application/x-ms-xbap;application/vnd.ms-xpsdocument;audio/midi;application/vnd.ibm.minipay;application/vnd.ibm.modcap;application/vnd.jcp.javame.midlet-rms;application/vnd.tmobile-ivetv;application/x-mobipocket-ebook;application/vnd.mobius.mbk;application/vnd.mobius.dis;application/vnd.mobius.plc;application/vnd.mobius.mqy;application/vnd.mobius.msl;application/vnd.mobius.txf;application/vnd.mobius.daf;tex/vnd.fly;application/vnd.mophun.certificate;application/vnd.mophun.application;video/mj2;audio/mpeg;video/vnd.mpegurl;video/mpeg;application/mp21;audio/mp4;video/mp4;application/mp4;application/vnd.apple.mpegurl;application/vnd.msician;application/vnd.muvee.style;application/xv+xml;application/vnd.nokia.n-gage.data;application/vnd.nokia.n-gage.symbian.install;application/x-dtbncx+xml;application/x-netcdf;application/vnd.neurolanguage.nlu;application/vnd.na;application/vnd.noblenet-directory;application/vnd.noblenet-sealer;application/vnd.noblenet-web;application/vnd.nokia.radio-preset;application/vnd.nokia.radio-presets;text/n3;application/vnd.novadigm.edm;application/vnd.novadim.edx;application/vnd.novadigm.ext;application/vnd.flographit;audio/vnd.nuera.ecelp4800;audio/vnd.nuera.ecelp7470;audio/vnd.nuera.ecelp9600;application/oda;application/ogg;audio/ogg;video/ogg;application/vnd.oma.dd2+xml;applicatin/vnd.oasis.opendocument.text-web;application/oebps-package+xml;application/vnd.intu.qbo;application/vnd.openofficeorg.extension;application/vnd.yamaha.openscoreformat;audio/webm;video/webm;application/vnd.oasis.opendocument.char;application/vnd.oasis.opendocument.chart-template;application/vnd.oasis.opendocument.database;application/vnd.oasis.opendocument.formula;application/vnd.oasis.opendocument.formula-template;application/vnd.oasis.opendocument.grapics;application/vnd.oasis.opendocument.graphics-template;application/vnd.oasis.opendocument.image;application/vnd.oasis.opendocument.image-template;application/vnd.oasis.opendocument.presentation;application/vnd.oasis.opendocumen.presentation-template;application/vnd.oasis.opendocument.spreadsheet;application/vnd.oasis.opendocument.spreadsheet-template;application/vnd.oasis.opendocument.text;application/vnd.oasis.opendocument.text-master;application/vnd.asis.opendocument.text-template;image/ktx;application/vnd.sun.xml.calc;application/vnd.sun.xml.calc.template;application/vnd.sun.xml.draw;application/vnd.sun.xml.draw.template;application/vnd.sun.xml.impress;application/vnd.sun.xl.impress.template;application/vnd.sun.xml.math;application/vnd.sun.xml.writer;application/vnd.sun.xml.writer.global;application/vnd.sun.xml.writer.template;application/x-font-otf;application/vnd.yamaha.openscoreformat.osfpvg+xml;application/vnd.osgi.dp;application/vnd.palm;text/x-pascal;application/vnd.pawaafile;application/vnd.hp-pclxl;application/vnd.picsel;image/x-pcx;image/vnd.adobe.photoshop;application/pics-rules;image/x-pict;application/x-chat;aplication/pkcs10;application/x-pkcs12;application/pkcs7-mime;application/pkcs7-signature;application/x-pkcs7-certreqresp;application/x-pkcs7-certificates;application/pkcs8;application/vnd.pocketlearn;image/x-portable-anymap;image/-portable-bitmap;application/x-font-pcf;application/font-tdpfr;application/x-chess-pgn;image/x-portable-graymap;image/png;image/x-portable-pixmap;application/pskc+xml;application/vnd.ctc-posml;application/postscript;application/xfont-type1;application/vnd.powerbuilder6;application/pgp-encrypted;application/pgp-signature;application/vnd.previewsystems.box;application/vnd.pvi.ptid1;application/pls+xml;application/vnd.pg.format;application/vnd.pg.osasli;tex/prs.lines.tag;application/x-font-linux-psf;application/vnd.publishare-delta-tree;application/vnd.pmi.widget;application/vnd.quark.quarkxpress;application/vnd.epson.esf;application/vnd.epson.msf;application/vnd.epson.ssf;applicaton/vnd.epson.quickanime;application/vnd.intu.qfx;video/quicktime;application/x-rar-compressed;audio/x-pn-realaudio;audio/x-pn-realaudio-plugin;application/rsd+xml;application/vnd.rn-realmedia;application/vnd.realvnc.bed;applicatin/vnd.recordare.musicxml;application/vnd.recordare.musicxml+xml;application/relax-ng-compact-syntax;application/vnd.data-vision.rdz;application/rdf+xml;application/vnd.cloanto.rp9;application/vnd.jisp;application/rtf;text/richtex;application/vnd.route66.link66+xml;application/rss+xml;application/shf+xml;application/vnd.sailingtracker.track;image/svg+xml;application/vnd.sus-calendar;application/sru+xml;application/set-payment-initiation;application/set-reistration-initiation;application/vnd.sema;application/vnd.semd;application/vnd.semf;application/vnd.seemail;application/x-font-snf;application/scvp-vp-request;application/scvp-vp-response;application/scvp-cv-request;application/svp-cv-response;application/sdp;text/x-setext;video/x-sgi-movie;application/vnd.shana.informed.formdata;application/vnd.shana.informed.formtemplate;application/vnd.shana.informed.interchange;application/vnd.shana.informed.package;application/thraud+xml;application/x-shar;image/x-rgb;application/vnd.epson.salt;application/vnd.accpac.simply.aso;application/vnd.accpac.simply.imp;application/vnd.simtech-mindmapper;application/vnd.commonspace;application/vnd.ymaha.smaf-audio;application/vnd.smaf;application/vnd.yamaha.smaf-phrase;application/vnd.smart.teacher;application/vnd.svd;application/sparql-query;application/sparql-results+xml;application/srgs;application/srgs+xml;application/sml+xml;application/vnd.koan;text/sgml;application/vnd.stardivision.calc;application/vnd.stardivision.draw;application/vnd.stardivision.impress;application/vnd.stardivision.math;application/vnd.stardivision.writer;application/vnd.tardivision.writer-global;application/vnd.stepmania.stepchart;application/x-stuffit;application/x-stuffitx;application/vnd.solent.sdkm+xml;application/vnd.olpc-sugar;audio/basic;application/vnd.wqd;application/vnd.symbian.install;application/smil+xml;application/vnd.syncml+xml;application/vnd.syncml.dm+wbxml;application/vnd.syncml.dm+xml;application/x-sv4cpio;application/x-sv4crc;application/sbml+xml;text/tab-separated-values;image/tiff;application/vnd.to.intent-module-archive;application/x-tar;application/x-tcl;application/x-tex;application/x-tex-tfm;application/tei+xml;text/plain;application/vnd.spotfire.dxp;application/vnd.spotfire.sfs;application/timestamped-data;applicationvnd.trid.tpt;application/vnd.triscape.mxs;text/troff;application/vnd.trueapp;application/x-font-ttf;text/turtle;application/vnd.umajin;application/vnd.uoml+xml;application/vnd.unity;application/vnd.ufdl;text/uri-list;application/nd.uiq.theme;application/x-ustar;text/x-uuencode;text/x-vcalendar;text/x-vcard;application/x-cdlink;application/vnd.vsf;model/vrml;application/vnd.vcx;model/vnd.mts;model/vnd.vtu;application/vnd.visionary;video/vnd.vivo;applicatin/ccxml+xml,;application/voicexml+xml;application/x-wais-source;application/vnd.wap.wbxml;image/vnd.wap.wbmp;audio/x-wav;application/davmount+xml;application/x-font-woff;application/wspolicy+xml;image/webp;application/vnd.webturb;application/widget;application/winhlp;text/vnd.wap.wml;text/vnd.wap.wmlscript;application/vnd.wap.wmlscriptc;application/vnd.wordperfect;application/vnd.wt.stf;application/wsdl+xml;image/x-xbitmap;image/x-xpixmap;image/x-xwindowump;application/x-x509-ca-cert;application/x-xfig;application/xhtml+xml;application/xml;application/xcap-diff+xml;application/xenc+xml;application/patch-ops-error+xml;application/resource-lists+xml;application/rls-services+xml;aplication/resource-lists-diff+xml;application/xslt+xml;application/xop+xml;application/x-xpinstall;application/xspf+xml;application/vnd.mozilla.xul+xml;chemical/x-xyz;text/yaml;application/yang;application/yin+xml;application/vnd.ul;application/zip;application/vnd.handheld-entertainment+xml;application/vnd.zzazz.deck+xml")
            fp.update_preferences()
            return fp.path
        else:
            print("Not able set firefox profile because browserName is : {} \n Provide browserName as firefox".format(
                browserName))

    @keyword
    def create_chrome_webdriver_and_enable_download_directory(self, browsername, download_dir, headless_mode='No'):
        """Used to create chrome webdriver with the options to set the default download directory and to launch in headless mode.
        By default the screen size for the browser is set to maximize.

        browsername : Browser name to be specified as chrome.
        download_dir : Directory of the path where the user wants to set the default download path.
                       Please see the below examples and give the directory path with two backslashes (\\\\)
        headless_mode : This enables user to launch the chrome in headless mode if set to yes,
                        By default headless mode is disabled. Please see the example 2.
        Example :
        1. To launch an url in UI mode with download directory set.
        | Create Chrome Webdriver And Enable Download Directory  |  chrome  |  D:\\\\Dummy_Download  |
        | Go To  |  https://www.finastra.com

        2. To launch an url in headless mode with download directory set.
        | Create Chrome Webdriver And Enable Download Directory  |  chrome  |  D:\\\\Dummy_Download  | headless_mode=yes
        | Go To  |  https://www.finastra.com
        """

        if browsername.lower() in ['chrome', 'googlechrome', 'gc']:
            chrome_options = Options()
            if headless_mode.lower() in ['yes']:
                chrome_options.add_experimental_option("prefs", {"download.default_directory": download_dir,
                                                                 "download.prompt_for_download": False})
                chrome_options.add_argument("--headless")
                chrome_options.add_argument('window-size=1920x1080')
                self.create_webdriver('Chrome', chrome_options=chrome_options)
                self._current_browser().command_executor._commands["send_command"] = (
                    "POST", '/session/$sessionId/chromium/send_command')
                params = {'cmd': 'Page.setDownloadBehavior',
                          'params': {'behavior': 'allow', 'downloadPath': download_dir}}
                self._current_browser().execute("send_command", params)
            else:
                prefs = {"download.default_directory": download_dir, 'safebrowsing.enabled': True}
                chrome_options.add_experimental_option("prefs", prefs)
                chrome_options.add_argument('--start-maximized')
                self.create_webdriver('Chrome', chrome_options=chrome_options)
        else:
            print("Not able create chrome webdriver because browserName is : {} \n Provide browserName as chrome".format(
                browsername))

    @keyword
    def press_shift_and_click_element(self, locator):
        """Used to hold shift key and Click on the element specified by 'locator'

        Arguments : 'locator' is the element locator

        Example :
        | Press Shift And Click Element | ${elementLocator} |

        """
        MoreOptions = self._element_find(locator, True, True)
        if MoreOptions:
            ActionChains(self._current_browser()).key_down(Keys.SHIFT).click(MoreOptions).key_up(
            Keys.SHIFT).perform()
        else:
            self.failure_occurred()
            raise AssertionError(" Element Not Found {} !! ".format(locator))

    @keyword
    def get_css_property(self, locator, propertyName, colorCode='HEX'):
        """ Used to get the css value of any element.If that element is not present then it will raise an Error

         Arguments : 'locator' is the element locator
                    'propertyName' is the css property name
                    'colorCode' is the code in which user wants the color or backgroud-color property. It can be either HEX or RGB.
                    This argument is case-insensetive.
                    This argument is applicable only for 'color' and 'background-color' propertyName.
                    Presently this keyword only supports 'HEX' and 'RGB' color code.

        Return: This keyword returns the value of the css-property mentioned in the argument propertyName

         Example :
        ${var1}    Get Css Property    ${test_login}    background-color    hex
        ${var2}    Get Css Property    ${test_login}    background
        ${var1}    Get Css Property    ${test_login}    margin-top
        ${var2}    Get Css Property    ${test_login}    background-color    rgb
        """
        web_element = self.get_webelement(locator)

        if web_element:
            strValue = web_element.value_of_css_property(propertyName)
            if strValue:
                print("Value of {} is {}".format(propertyName, strValue))
                a = ['color', 'background-color']
                if propertyName in a:
                    if colorCode.upper() == 'HEX':
                        return self._rgb_to_hex(strValue)
                    elif colorCode.upper() == 'RGB':
                        return self._hex_to_rgb(strValue)
                    else:
                        print("Color code {} is not supported".format(colorCode))
                        return strValue
                else:
                    return strValue
            else:
                self.failure_occurred()
                raise AssertionError(propertyName + " property Not Found")
        else:
            self.failure_occurred()
            raise AssertionError(locator + " Element Not Found")

    def _is_rgb(self, s):
        try:
            if (s.startswith('rgba(') and s.endswith(')')):
                return True
        except ValueError:
            return False

    def _rgb_to_hex(self, str1):
        if (self._is_rgb(str1)):
            r, g, b, a = str1.replace('rgba(', "").replace(")", "").split(",")
            hex = "#{:02x}{:02x}{:02x}".format(int(r), int(g), int(b))
            return hex
        else:
            return str1

    def _is_hex(self, s):
        try:
            int(s.replace('#', ''), 16)
            return True
        except ValueError:
            return False

    def _hex_to_rgb(self, hexcode):
        if (self._is_hex(hexcode)):
            rgb = tuple(map(ord, hexcode[1:].decode('hex')))
            return rgb
        else:
            return hexcode

    def _wait_for(self,condition_function):
        start_time = time.time()
        while time.time() < start_time + 10:
            if condition_function():
                return True
            else:
                time.sleep(0.1)
        raise Exception(
            'Timeout waiting for {}'.format(condition_function.__name__)
        )

    def _page_has_loaded(self):
        page_state = self._current_browser().execute_script(
            'return document.readyState;'
        )
        return page_state == 'complete'

    @keyword
    def download_file(self):
        """Downloads a file by clicking on a particular button.
            Supported browsers are Edge, IE, Chrome and firefox.
                Arguments:
                - No arguments are required. It automatically downloads the file based on the currently invoked browser.

                Pre-conditions:
                In order to use this keyword for Chrome, browser should be invoked by using the below resource Keyword under Selenium > Common > Browser > browser_keywords.robot

                Open Chrome Browser by setting download path    ${URL}    ${FolderLocation}

                In order to use this keyword for firefox, browser should be invoked by using the below resource Keyword Selenium > Common > Browser > browser_keywords.robot

                Open Firefox Browser by setting download path   ${URL}    ${FolderLocation}

                User can set the default download directory in IE by the below mentioned steps.
                    Open Internet Explorer, select the Tools button, and then select View downloads.
                    In the View Downloads dialog box, select Options in the lower left hand corner.
                    Choose a different default download location by selecting Browse, and then select OK when you're done.

                User can set the default download directory in Edge by the below mentioned steps.
                    Launch Edge from your Start menu, desktop, or taskbar.
                    Click the More button - it is located near the top right corner of the window and looks like "..."
                    Click Settings.
                    Click View advanced settings. You might have to scroll down a bit to find it.
                    Click Change.
                    Click a new folder where you'd like to save files.
                    Click Select Folder.
                    Click the More button to close the menu.

                Example:

                #After step to click element on the button that downloads the file. Next step should be this keyword.
                | Download File|
                """

        caps = self._current_browser().capabilities
        browserName = caps['browserName']
        k = KeyboardLibrary()

        try:
            self._wait_for(self._page_has_loaded)
        except:
            pass

        if browserName in ['internet explorer', 'ie','MicrosoftEdge']:
            time.sleep(4)
            k.native_type("{TAB}")
            time.sleep(1)
            k.native_type("{TAB}")
            time.sleep(1)
            k.native_type("{TAB}")
            time.sleep(1)
            k.native_type("{ENTER}")
            time.sleep(3)
        elif browserName in ['firefox', 'ff']:
            time.sleep(4)
            k.native_type("{ENTER}")
            time.sleep(2)
        elif browserName in ['googlechrome', 'gc', 'chrome']:
            pass
        else:
            raise AssertionError("browser {} not supported.".format(browserName))

    @keyword
    def download_pdf(self, filename=None):
        """Downloads a pdf by clicking on a particular link that opens the pdf.
            Supported browsers are IE, Chrome and firefox.
                Arguments:
                - For chrome and firefox browser, no arguments are required. It automatically downloads the pdf file
                    based on the currently invoked browser.
                - filename: Needed only for IE browser. This argument specifies the full name of the file in which user
                            wants to save the pdf, along with it's location.
                Pre-conditions:

                In order to use this keyword for Chrome, browser should be invoked by using the below resource Keyword Selenium > Common > Browser > browser_keywords.robot

                 Open Chrome Browser by setting download path    ${URL}    ${FolderLocation}

                In order to use this keyword for firefox, browser should be invoked by using the below resource Keyword Selenium > Common > Browser > browser_keywords.robot

                Open Firefox Browser by setting download path   ${URL}    ${FolderLocation}

                User can set the default download directory in IE by the below mentioned steps.
                    Open Internet Explorer, select the Tools button, and then select View downloads.
                    In the View Downloads dialog box, select Options in the lower left hand corner.
                    Choose a different default download location by selecting Browse, and then select OK when you're done.

                Example:
                #After step to click element on the button that downloads/opens the pdf. Next step should be this keyword.
                #for chrome and firefox browser.
                | Download Pdf|
                #for IE browser
                Download Pdf    D:\\Rough\\ie\\test13.pdf
                """
        caps = self._current_browser().capabilities
        browserName = caps['browserName']
        k = KeyboardLibrary()
        if browserName in ['internet explorer', 'ie']:
            if filename is None:
                raise AssertionError("value for filename argument is required.")
            folder_name = filename.replace(os.path.basename(filename), "")
            if not os.path.exists(folder_name):
                raise AssertionError("Folder {} doesnot exist.".format(folder_name))
            i = 0
            file = filename
            while os.path.isfile(file):
                i += 1
                file = filename.replace(os.path.splitext(filename)[1], "") + "_{}".format(i) + \
                       os.path.splitext(filename)[1]
            time.sleep(4)
            x = pyautogui.size()[0]
            y = pyautogui.size()[1]
            pyautogui.click(x / 2, y / 2)
            time.sleep(2)
            k.press_combination("KEY.CTRL", "KEY.SHIFT", "KEY.S")
            time.sleep(2)
            k.native_type(file)
            time.sleep(2)
            k.native_type("{ENTER}")
            time.sleep(2)
            print(file + " saved.")
        elif browserName in ['firefox', 'ff']:
            self.wait_until_element_is_visible("//button[@id='download']",20)
            element = self._element_find("//button[@id='download']", True, True)
            element.click()
            time.sleep(2)
            k.native_type("{ENTER}")
        elif browserName in ['googlechrome', 'gc', 'chrome']:
            pass
        else:
            raise AssertionError("browser {} not supported.".format(browserName))

    @keyword
    def print_pdf(self):

        """Prints a pdf by clicking on a particular link that opens the pdf.
            Supported browsers are IE, Chrome and firefox.
                        Arguments:
                        - No arguments are required. It automatically prints the pdf file based on the currently invoked browser.

                        Pre-conditions:

                        Example:
                        #After step to click element on the button that downloads/opens the pdf. Next step should be this keyword.
                        | Print Pdf|

                        """
        caps = self._current_browser().capabilities
        browserName = caps['browserName']
        k = KeyboardLibrary()
        if browserName in ['internet explorer', 'ie']:
            time.sleep(4)
            x = pyautogui.size()[0]
            y = pyautogui.size()[1]
            pyautogui.click(x / 2, y / 2)
            time.sleep(2)
            k.press_combination("KEY.CTRL", "KEY.P")
            time.sleep(3)
            k.native_type("{ENTER}")
            time.sleep(4)
        elif browserName in ['Chrome', 'chrome', 'Firefox', 'ff', 'firefox']:
            time.sleep(2)
            k.press_combination("KEY.CTRL", "KEY.P")
            time.sleep(2)
            k.native_type("{ENTER}")
            time.sleep(4)
        elif browserName in ['googlechrome', 'gc', 'chrome']:
            pass
        else:
            raise AssertionError("browser {} not supported.".format(browserName))

    def _set_attribute_javascript(self, locator, attribute, value, index):

        web_elements = self.get_webelements(locator)
        if len(web_elements) == 0:
            raise AssertionError("Element {} not found.".format(locator))
        element = web_elements[int(index) - 1]
        script = "arguments[0].setAttribute('{}', '{}')".format(attribute, value)
        self._current_browser().execute_script(script, element)

    @keyword
    def set_attribute_javascript(self, locator, index, *params):

        """ Set's the attribute of an element specified by 'locator'. The element could be either visible or hidden.

         |Arguments|

         'locator' = is the element locator \n

         'index' = Used to select a specific element from multiple occurrence of elements specified by 'locator'.\n

         'params' = Used to pass in the format attribute=value. It takes multiple attributes separated by Tab space. For example: class=abc \n


        |Example|
        | ***TestCases*** |

         1. To set single attribute of an element specified by 'locator' with index as 1 (selects the first element)
         | Set Attribute Javascript | //span[@class='glyphicon glyphicon-record']/../div | 1 | class=abc |

         2. To set single attribute of an element specified by 'locator' with index as 3 (selects the third element)
         | Set Attribute Javascript | //span[@class='glyphicon glyphicon-record']/../div | 3 | class=abc |

         3. To set multiple attribute of an element specified by 'locator' with index as 2 (selects the second element)
         | Set Attribute Javascript | //span[@class='glyphicon glyphicon-record']/../div | 2 | class=${EMPTY} | aria-hidden=false |
        """
        if int(index) < 1:
            raise AssertionError("Index should be greater than or equal to 1.")
        for param in params:
            self._set_attribute_javascript(locator, param.split("=")[0], param.split("=")[1], index)

    @keyword()
    def js_click(self, locator):

        """ Used to click an element using javascriptexecutor in Selenium. This keyword can be used if "Click Element"
         keyword is unable to click on the desired element.

        |Arguments|

                'locator' = is the element locator \n

        |Example|

            Js Click    //input[@name="btnK"]
        """
        element = self._element_find(locator, True, True)
        script = 'arguments[0].click()'
        self._current_browser().execute_script(script, element)

    @keyword
    def press_down_arrow_and_enter_key(self):
        """Used to press windows keyboard's Arrow Down key followed by Enter key on the application.

        Example :
        | Press Down Arrow And Enter Key |

        """
        try:
            ActionChains(self._current_browser()).key_down(Keys.ARROW_DOWN).key_down(Keys.ENTER).perform()
        except:
            self.failure_occurred()
            raise AssertionError("Could not perform the action")

    @keyword
    def get_element_attribute(self, locator, attribute=None):
        """Returns value of ``attribute`` from element ``locator``.

        See the `Locating elements` section for details about the locator
        syntax.

        Example:
        | ${id}= | `Get Element Attribute` | css:h1 | id |
        | ${id}= | `Get Element Attribute` | css:h1@id |

        Passing attribute name as part of the ``locator`` or The explicit ``attribute`` argument
        both conditions are handled.
        """
        if attribute is None:
            locator, attribute = locator.rsplit('@', 1)
        return self.find_element(locator).get_attribute(attribute)

    @keyword
    def capture_full_page_screenshot(self, screenshot_name=None):
        """Takes full page screenshot of the current page and embeds it into log file.

        |Arguments|

        `screenshot_name` = Argument specifies the name of the file to write the screenshot into.

        Example :
        | Capture Full Page Screenshot |
        | Capture Full Page Screenshot | app_screenshot
        """
        screenshot_filename = 'full-page-selenium-screenshot-{index}.png'
        if screenshot_name:
            screenshot_filename = str(screenshot_name) + '-{index}.png'
        js = 'return Math.max( document.body.scrollHeight, document.body.offsetHeight,  document.documentElement.clientHeight,  document.documentElement.scrollHeight,  document.documentElement.offsetHeight);'  # get the maximum height of the page
        scrollheight = self._current_browser().execute_script(js)
        offset = 0
        try:
            while offset < scrollheight:
                self._current_browser().execute_script("window.scrollTo(0, %s);" % offset)  # scroll to size of page
                img = Image.open(BytesIO(self._current_browser().get_screenshot_as_png()))  # Create image
                offset += img.size[1]  # Increment the height of the page
                self.capture_page_screenshot(filename=screenshot_filename)
        except Exception as e:
            raise AssertionError(e)
