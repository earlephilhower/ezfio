#!/usr/bin/python

# ezfio 1.0
# earle.philhower.iii@hgst.com
#
# ------------------------------------------------------------------------
# ezfio is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 2 of the License, or
# (at your option) any later version.
# 
# ezfio is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with ezfio.  If not, see <http://www.gnu.org/licenses/>.
# ------------------------------------------------------------------------
#
# Usage:   ./ezfio.py -d </dev/node> [-u <100..1>]
# Example: ./ezfio.py -d /dev/nvme0n1 -u 100
# 
# This script requires root privileges so must be run as "root" or
# via "sudo ./ezfio.py"
#
# Please be sure to have FIO installed, or you will be prompted to install
# and re-run the script.

import argparse
import base64
import datetime
import os
import platform
import pwd
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import zipfile

def AppendFile(text, filename):
    """Equivalent to >> in BASH, append a line to a text file."""
    with open(filename, "a") as f:
        f.write(text)
        f.write("\n")

def Run(cmd):
    """Run a cmd[], return the exit code, stdout, and stderr."""
    proc=subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out = proc.stdout.read()
    err = proc.stderr.read()
    code = proc.wait()
    return code, out, err

def CheckAdmin():
    """Check that we have root privileges for disk access, abort if not."""
    if ( os.geteuid() != 0 ):
        sys.stderr.write("Root privileges are required for low-level disk ")
        sys.stderr.write("access.\nPlease restart this script as root ")
        sys.stderr.write("(sudo) to continue.\n")
        sys.exit(1)

def FindFIO():
    """Try the path and the CWD for a FIO executable, return path or exit."""
    # Determine if FIO is in path or CWD
    try:
        ret, out, err = Run(["fio", "-v"])
        if (ret == 0):
            return "fio"
    except:
        try:
            ret, out, err = Run(['./fio', '-v'])
            if (ret == 0):
                return "./fio"
        except:
            sys.stderr.write("FIO is required to run IO tests.\n")
            sys.stderr.write("The latest versions can be found at ")
            sys.stderr.write("https://github.com/axboe/fio.\n")
            sys.exit(1)

def CheckFIOVersion():
    """Check that we have a version of FIO installed that we can use."""
    global fio, fioVerString
    code, out, err = Run( [fio, '--version'] )
    try:
        fioVerString = out.split('\n')[0].rstrip()
        ver = out.split('\n')[0].rstrip().split('-')[1].split('.')[0]
        if int(ver) < 2:
            sys.stderr.write("ERROR: FIO version " + ver + " unsupported, ")
            sys.stderr.write("version 2.0 or later required.  Exiting.\n")
            sys.exit(2)
    except:
        sys.stderr.write("ERROR: Unable to determine version of fio ")
        sys.stderr.write("installed.  Exiting.\n")
        sys.exit(2)


def ParseArgs():
    """Parse command line options into globals."""
    global physDrive, utilization, yes
    parser = argparse.ArgumentParser(
                 formatter_class=argparse.RawDescriptionHelpFormatter,
    description="A tool to easily run FIO to benchmark sustained " \
                "performance of NVME\nand other types of SSD.",
    epilog="""
Requirements:\n
* Root access (log in as root, or sudo {prog})
* No filesytems or data on target device
* FIO IO tester (available https://github.com/axboe/fio)
* sdparm to identify the NVME device and serial number

WARNING: All data on the target device will be DESTROYED by this test.""")
    parser.add_argument("--drive", "-d", dest = "physDrive",
        help="Device to test (ex: /dev/nvme0n1)", required=True)
    parser.add_argument("--utilization", "-u", dest="utilization",
        help="Amount of drive to test (in percent), 1...100", default="100",
        type=int, required=False)
    parser.add_argument("--yes", dest="yes", action='store_true',
        help="Skip the final warning prompt (for scripted tests)",
        required=False)
    args = parser.parse_args()

    physDrive = args.physDrive
    utilization = args.utilization
    yes = args.yes
    if (utilization < 1) or (utilization > 100):
        print "ERROR:  Utilization must be between 1...100"
        parser.print_help()
        sys.exit(1)
    # Sanity check that the selected drive is not mounted by parsing mounts
    # This is not guaranteed to catch all as there's just too many different
    # naming conventions out there.  Let's cover simple HDD/SSD/NVME patterns
    if ( re.match('.*p?[1-9][0-9]*$', physDrive) and
         not re.match('.*/nvme[0-9]+n[1-9][0-9]*$', physDrive) ):
        pdispart = True
    else:
        pdispart = False
    hit = ""
    with open("/proc/mounts", "r") as f:
        mounts = f.readlines()
    for l in mounts:
        dev = l.split()[0]
        mnt = l.split()[1]
        if dev == physDrive:
            hit = dev + " on " + mnt # Obvious exact match
        if pdispart:
            chkdev = dev
        else:
            # /dev/sdp# is special case, don't remove the "p"
            if re.match('^/dev/sdp.*$', dev):
                chkdev = re.sub('[1-9][0-9]*$', '', dev)
            else:
                # Need to see if mounted partition is on a raw device being tested
                chkdev = re.sub('p?[1-9][0-9]*$', '', dev)
        if chkdev == physDrive: hit = dev + " on " + mnt
    if hit != "" :
        print "ERROR:  Mounted volume '" + str(hit) + "' is on same device",
        print "as tested device '" + str(physDrive) + "'.  ABORTING."
        sys.exit(2)


def CollectSystemInfo():
    """Collect some OS and CPU information."""
    global cpu, cpuCores, cpuFreqMHz, uname
    uname = " ".join(platform.uname())
    code, cpuinfo, err = Run(['cat', '/proc/cpuinfo'])
    cpuinfo = cpuinfo.split("\n")
    code, dmidecode, err = Run(['dmidecode', '--type', 'processor'])
    if 'ppc64' in uname:
        # Implement grep and sed in Python...
        cpu = filter(lambda x:re.search(r'model', x), cpuinfo)[0].split(': ')[1].replace('(R)','').replace('(TM)','')
        cpuCores = len(filter(lambda x:re.search('processor', x), cpuinfo))
        try:
            cpuFreqMHz = int(round(float(filter(lambda x: re.search('Current Speed', x), dmidecode.split("\n"))[0].rstrip().lstrip().split(" ")[2])))
        except:
            cpuFreqMHz = int(round(float(filter(lambda x:re.search('clock', x), cpuinfo)[0].split(': ')[1][:-3])))
    else:
        cpu = filter(lambda x:re.search(r'model name', x), cpuinfo)[0].split(': ')[1].replace('(R)','').replace('(TM)','')
        cpuCores = len(filter(lambda x:re.search('model name', x), cpuinfo))
        try:
            cpuFreqMHz = int(round(float(filter(lambda x: re.search('Current Speed', x), dmidecode.split("\n"))[0].rstrip().lstrip().split(" ")[2])))
        except:
            cpuFreqMHz = int(round(float(filter(lambda x:re.search('cpu MHz', x), cpuinfo)[0].split(': ')[1])))

def VerifyContinue():
    """User's last chance to abort the test.  Exit if they don't agree."""
    if not yes:
        print "-" * 75
        print "WARNING! " * 9
        print "THIS TEST WILL DESTROY ANY DATA AND FILESYSTEMS ON ",
        print physDrive +"\n"
        cont = raw_input("Please type the word \"yes\" and hit return to " +
                         "continue, or anything else to abort.\n")
        print "-" * 75 + "\n"
        if cont != "yes":
            print "Performance test aborted, drive is untouched.\n"
            sys.exit(1)


def CollectDriveInfo():
    """Get important device information, exit if not possible."""
    global physDriveGiB, physDriveGB, physDriveBase, testcapacity
    global model, serial, physDrive
    # We absolutely need this information
    try:
        physDriveBase = os.path.basename(physDrive)
        code, physDriveBytes, err=Run(['blockdev', '--getsize64', physDrive])
        if code != 0:
            raise Exception("Can't get drive size for " + physDrive)
        physDriveGB = (long(physDriveBytes))/(1000 * 1000 * 1000)
        physDriveGiB = (long(physDriveBytes))/(1024 * 1024 * 1024)
        testcapacity = (physDriveGiB * utilization) / 100
    except:
        print "ERROR: Can't get '" + physDrive + "' size. ",
        print "Incorrect device name?"
        sys.exit(1)
    # These are nice to have, but we can run without it
    model = "UNKNOWN"
    serial = "UNKNOWN"
    try:
        sdparmcmd = ['sdparm', '--page', 'sn', '--inquiry', '--long',
                     physDrive]
        code, sdparm, err = Run(sdparmcmd)
        lines = sdparm.split("\n")
        if len(lines) == 4:
            model=re.sub("\s+", " ", lines[0].split(":")[1].lstrip().rstrip())
            serial = re.sub("\s+", " ", lines[2].lstrip().rstrip())
        else:
            print "Unable to identify drive using sdparm. Continuing."
    except:
        print "Install sdparm to allow model/serial extraction. Continuing."



def SetupFiles():
    """Set up names for all output/input files, place headers on CSVs."""
    global ds, pwd, details, testcsv, timeseriescsv, odssrc, odsdest
    global physDriveBase, fioVerString

    def CSVInfoHeader(f):
        """Headers to the CSV file (ending up in the ODS at the test end)."""
        global physDrive, model, serial, physDriveGiB, testcapacity
        global cpu, cpuCores, cpuFreqMHz, uname
        AppendFile("Drive," + str(physDrive), f)
        AppendFile("Model," + str(model), f)
        AppendFile("Serial," + str(serial), f)
        AppendFile("AvailCapacity," + str(physDriveGiB) + ",GiB", f)
        AppendFile("TestedCapacity," + str(testcapacity) + ",GiB", f)
        AppendFile("CPU," + str(cpu), f)
        AppendFile("Cores," + str(cpuCores), f)
        AppendFile("Frequency," + str(cpuFreqMHz), f)
        AppendFile("OS," + str(uname), f)
        AppendFile("FIOVersion," + str(fioVerString), f)

    # Datestamp for run output files
    ds = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # The unique suffix we generate for all output files
    suffix  = str(physDriveGB) + "GB_" + str(cpuCores) + "cores_"
    suffix += str(cpuFreqMHz) + "MHz_" + physDriveBase + "_"
    suffix += socket.gethostname() + "_" + ds

    pwd = os.getcwd()
    # The "details" directory contains the raw output of each FIO run
    details = pwd + "/details_" + suffix
    if os.path.exists(details):
        shutil.rmtree(details)
    os.mkdir(details)
    # Copy this script into it for posterity
    shutil.copyfile(__file__, details + "/" + os.path.basename(__file__) )

    # Files we're going to generate, encode some system info in the names
    # If the output files already exist, erase them
    testcsv = details + "/ezfio_tests_"+suffix+".csv"
    if os.path.exists(testcsv):
        os.unlink(testcsv)
    CSVInfoHeader(testcsv)
    AppendFile("Type,Write %,Block Size,Threads,Queue Depth/Thread,IOPS," +
               "Bandwidth (MB/s),Read Latency (us),Write Latency (us)",
               testcsv)
    timeseriescsv = details + "/ezfio_timeseries_"+suffix+".csv"
    if os.path.exists(timeseriescsv):
        os.unlink(timeseriescsv)
    CSVInfoHeader(timeseriescsv)
    AppendFile("IOPS", timeseriescsv) # Add IOPS header

    # ODS input and output files
    odssrc = pwd + "/original.ods"
    odsdest = pwd + "/ezfio_results_"+suffix+".ods"
    if os.path.exists(odsdest):
        os.unlink(odsdest)

class FIOError(Exception):
    """Exception generated when FIO returns a non-success value

    Attributes:
        cmdline -- The FIO command that was executed
        code    -- Error code FIO returned
        stderr  -- STDERR output from FIO
        stdout  -- STDOUT output from FIO
    """
    def __init__(self, cmdline, code, stderr, stdout):
        self.cmdline = cmdline
        self.code = code
        self.stderr = stderr
        self.stdout = stdout

def SequentialConditioning():
    """Sequentially fill the complete capacity of the drive once."""
    # Note that we can't use regular test runner because this test needs
    # to run for a specified # of bytes, not a specified # of seconds.
    cmdline = [fio, "--name=SeqCond", "--readwrite=write", "--bs=128k", 
               "--ioengine=libaio", "--iodepth=64", "--direct=1", 
               "--filename=" + physDrive, "--size=" + str(testcapacity) + "G",
               "--thread"]
    code, out, err = Run(cmdline)
    if code != 0:
        raise FIOError(" ".join(cmdline), code , err, out)
    else:
        return "DONE", "DONE", "DONE"

def RandomConditioning():
    """Randomly write entire device for the full capacity"""
    # Note that we can't use regular test runner because this test needs
    # to run for a specified # of bytes, not a specified # of seconds.
    cmdline = [fio, "--name=RandCond", "--readwrite=randwrite", "--bs=4k",
               "--invalidate=1", "--end_fsync=0", "--group_reporting",
               "--direct=1", "--filename=" + str(physDrive),
               "--size=" + str(testcapacity) + "G", "--ioengine=libaio",
               "--iodepth=256", "--norandommap", "--randrepeat=0", "--thread"]
    code, out, err = Run(cmdline)
    if code != 0:
        raise FIOError(" ".join(cmdline), code , err, out)
    else:
        return "DONE", "DONE", "DONE"

def RunTest(iops_log, seqrand, wmix, bs, threads, iodepth, runtime):
    """Runs the specified test, generates output CSV lines."""

    def IOStatThread(**kwargs):
        """Collect 1-second interval IOPS values to a CSV."""
        starttime = datetime.datetime.now()
        stoptime = starttime + datetime.timedelta(0, int(o['runtime']))
        statpath = "/sys/block/"+physDriveBase+"/stat"
        if not os.path.exists(statpath):
            base = re.sub("[0-9]+$", "", physDriveBase)
            statpath = "/sys/block/"+base+"/stat"
        with open(statpath, "r") as f:
            stat = f.read().rstrip().split()
            readstart = long(stat[0])
            writestart = long(stat[4])
        timeseries = open(timeseriescsv, "a")
        now = starttime
        while now < stoptime:
            time.sleep(1)
            with open(statpath, "r") as f:
                stat = f.read().rstrip().split()
                readend = long(stat[0])
                writeend = long(stat[4])
            iops = (readend - readstart) + (writeend - writestart)
            readstart = readend
            writestart = writeend
            timeseries.write(str(iops) + "\n")
            timeseries.flush()
            now = datetime.datetime.now()
        timeseries.close()

    # Output file names
    testfile  = str(details) + "/Test" + str(seqrand) + "_w" + str(wmix)
    testfile += "_bs" + str(bs) + "_threads" + str(threads) + "_iodepth"
    testfile += str(iodepth) + "_"+str(physDriveBase) + ".out"
    
    if seqrand == "Seq":
        rw = "rw"
    else:
        rw = "randrw" 

    if iops_log:
        o={}
        o['runtime'] = runtime
        iostat = threading.Thread(target=IOStatThread, kwargs=(o))
        iostat.start()

    cmdline = [fio, "--name=test", "--readwrite=" + str(rw),
               "--rwmixwrite=" + str(wmix), "--bs=" + str(bs),
               "--invalidate=1", "--end_fsync=0", "--group_reporting",
               "--direct=1", "--filename=" + str(physDrive),
               "--size=" + str(testcapacity) + "G", "--time_based",
               "--runtime=" + str(runtime), "--ioengine=libaio",
               "--numjobs=" + str(threads), "--iodepth=" + str(iodepth),
               "--norandommap", "--randrepeat=0", "--thread",
               "--output-format=terse", "--terse-version=3",
               "--exitall"]
    
    AppendFile(" ".join(cmdline), testfile)

    # There are some NVME drives with 4k physical and logical out there.
    # Check that we can actually do this size IO, OTW return 0 for all
    skiptest = False
    code, out, err = Run(['blockdev', '--getiomin', str(physDrive)])
    if code == 0: 
        iomin = int(out.split("\n")[0])
        if int(bs) < iomin: skiptest = True
    # Silently ignore failure to return min block size, FIO will fail and
    # we'll catch that a little later.
    if skiptest:
        code = 0
        out = "Test not run because block size " + str(bs)
        out += "below iominsize " + str(iomin) + "\n"
        out += "3;" + "0;" * 100 + "\n"  # Bogus 0-filled resulte line
        err = ""
    else:
        code, out, err = Run(cmdline)
    AppendFile("[STDOUT]", testfile)
    AppendFile(out, testfile)
    AppendFile("[STDERR]", testfile)
    AppendFile(err, testfile)

    # Make sure we had successful completion, else note and abort run
    if code != 0:
        AppendFile("ERROR", testcsv)
        raise FIOError(" ".join(cmdline), code, err, out)

    if iops_log:
        iostat.join()

    lines = out.split("\n")
    # Terse position varies, look for identification at line start "3;"
    resultsline = filter(lambda x:re.search(r'^3;', x), lines)
    results = resultsline[0].split(";")

    rdiops = float(results[7])
    wriops = float(results[48])
    rlat = float(results[39])
    wlat = float(results[80])
    iops = "{0:0.0f}".format( rdiops + wriops )
    mbps = "{0:0.2f}".format((float( (rdiops+wriops) * bs ) /
                                     ( 1024.0 * 1024.0 )))
    lat = "{0:0.1f}".format(max(rlat, wlat))

    AppendFile( ",".join((str(seqrand), str(wmix), str(bs), str(threads),
                          str(iodepth), str(iops), str(mbps), str(rlat),
                          str(wlat))), testcsv)
    return iops, mbps, lat

def DefineTests():
    """Generate the work list for the main worker into OC."""
    global oc
    # What we're shmoo-ing across
    bslist = (512, 1024, 2048, 4096, 8192, 16384, 32768, 65536, 131072)
    qdlist = (1, 2, 4, 8, 16, 32, 64, 128, 256)
    threadslist = (1, 2, 4, 8, 16, 32, 64, 128, 256)
    shorttime = 120 # Runtime of point tests
    longtime = 1200 # Runtime of long-running tests

    def AddTest( name, seqrand, writepct, blocksize, threads, qdperthread,
                 iops_log, runtime, desc, cmdline ):
        if threads != "":
            qd = int(threads) * int(qdperthread)
        else:
            qd = 0
        dat = {}
        dat['name'] = name
        dat['seqrand'] = seqrand
        dat['wmix'] = writepct
        dat['bs'] = blocksize
        dat['qd'] = qd
        dat['qdperthread'] = qdperthread
        dat['threads'] = threads
        dat['bw'] = ''
        dat['iops'] = ''
        dat['lat'] = ''
        dat['desc'] = desc
        dat['iops_log'] = iops_log
        dat['runtime'] = runtime
        dat['cmdline'] = cmdline
        oc.append(dat)

    def DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, desc,
                  iops_log, runtime):
        AddTest(testname, seqrand, wmix, bs, threads, iodepth, iops_log,
                runtime, desc, lambda o: {RunTest(o['iops_log'],
                                                  o['seqrand'], o['wmix'],
                                                  o['bs'], o['threads'],
                                                  o['qdperthread'],
                                                  o['runtime'])})

    def AddTestBSShmoo():
        AddTest(testname, 'Preparation', '', '', '', '', '', '', '',
                lambda o: {AppendFile(o['name'], testcsv)} )
        for bs in bslist:
            desc = testname + ", BS=" + str(bs)
            DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, desc,
                      iops_log, runtime)

    def AddTestQDShmoo():
        AddTest(testname, 'Preparation', '', '', '', '', '', '', '',
                lambda o: {AppendFile(o['name'], testcsv)} )
        for iodepth in qdlist:
            desc = testname + ", QD=" + str(iodepth)
            DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, desc,
                      iops_log, runtime)

    def AddTestThreadsShmoo():
        AddTest(testname, 'Preparation', '', '', '', '', '', '', '',
                lambda o: { AppendFile(o['name'], testcsv ) } )
        for threads in threadslist:
            desc = testname + ", Threads=" + str(threads)
            DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, desc,
                      iops_log, runtime)

    AddTest('Sequential Preconditioning', 'Preparation', '', '', '', '', '',
            '', '', lambda o: {} ) # Only for display on-screen
    AddTest('Sequential Preconditioning', 'Seq Pass 1', '100', '131072', '1',
            '256', False, '', 'Sequential Preconditioning Pass 1',
            lambda o: {SequentialConditioning()} )
    AddTest('Sequential Preconditioning', 'Seq Pass 2', '100', '131072', '1',
            '256', False, '', 'Sequential Preconditioning Pass 2',
            lambda o: {SequentialConditioning()} )

    testname = "Sustained Multi-Threaded Sequential Read Tests by Block Size"
    seqrand = "Seq"
    wmix=0
    threads=1
    runtime=shorttime
    iops_log=False
    iodepth=256
    AddTestBSShmoo()

    testname = "Sustained Multi-Threaded Random Read Tests by Block Size"
    seqrand = "Rand"
    wmix=0
    threads=16
    runtime=shorttime
    iops_log=False
    iodepth=16
    AddTestBSShmoo()

    testname = "Sequential Write Tests with Queue Depth=1 by Block Size"
    seqrand = "Seq"
    wmix=100
    threads=1
    runtime=shorttime
    iops_log=False
    iodepth=1
    AddTestBSShmoo()

    AddTest('Random Preconditioning', 'Preparation', '', '', '', '', '', '',
            '', lambda o: {} ) # Only for display on-screen
    AddTest('Random Preconditioning', 'Rand Pass 1', '100', '4096', '1',
            '256', False, '', 'Random Preconditioning',
            lambda o: {RandomConditioning()} )
    AddTest('Random Preconditioning', 'Rand Pass 2', '100', '4096', '1',
            '256', False, '', 'Random Preconditioning',
            lambda o: {RandomConditioning()} )

    testname = "Sustained 4KB Random Read Tests by Number of Threads"
    seqrand = "Rand"
    wmix=0
    bs=4096
    runtime=shorttime
    iops_log=False
    iodepth=1
    AddTestThreadsShmoo()

    testname = "Sustained 4KB Random mixed 30% Write Tests by Threads"
    seqrand = "Rand"
    wmix=30
    bs=4096
    runtime=shorttime
    iops_log=False
    iodepth=1
    AddTestThreadsShmoo()

    testname = "Sustained Perf Stability Test - 4KB Random 30% Write"
    AddTest(testname, 'Preparation', '', '', '', '', '', '', '',
            lambda o: {AppendFile(o['name'], testcsv)} )
    seqrand = "Rand"
    wmix=30
    bs=4096
    runtime=longtime
    iops_log=True
    iodepth=1
    threads=256
    DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, testname,
              iops_log, runtime)

    testname = "Sustained 4KB Random Write Tests by Number of Threads"
    seqrand = "Rand"
    wmix=100
    bs=4096
    runtime=shorttime
    iops_log=False
    iodepth=1
    AddTestThreadsShmoo()

    testname = "Sustained Multi-Threaded Random Write Tests by Block Size"
    seqrand = "Rand"
    wmix=100
    runtime=shorttime
    iops_log=False
    iodepth=16
    threads=16
    AddTestBSShmoo()


def RunAllTests():
    """Iterate through the OC work queue and run each job, show progress."""
    global ret_iops, ret_mbps, ret_lat

    # Determine some column widths to make format specifiers
    maxlen = 0
    for o in oc:
        maxlen = max(maxlen, len(o['desc']))
    descfmt = "{0:" + str(maxlen) + "}"
    resfmt = "{1: >8} {2: >9} {3: >8}"
    fmtstr = descfmt + " " + resfmt

    def JobWrapper(**kwargs):
        """Thread wrapper to store return values for parent to read later."""
        global ret_iops, ret_mbps, ret_lat, oc
        # Until we know it's succeeded, we're in error
        ret_iops = "ERROR"
        ret_mbps = "ERROR"
        ret_lat = "ERROR"
        try:
            val = o['cmdline'](o)
            ret_iops = list(val)[0][0]
            ret_mbps = list(val)[0][1]
            ret_lat = list(val)[0][2]
        except FIOError as e:
            print "\nFIO Error!\n" + e.cmdline + "\nSTDOUT:\n" + e.stdout
            print "STDERR:\n" + e.stderr
            raise
        except:
            print "\nUnexpected error while running FIO job."
            raise
        
    print "*" * len(fmtstr.format("", "", "", ""))
    print "ezFio test parameters:\n"

    fmtinfo="{0: >20}: {1}"
    print fmtinfo.format("Drive", str(physDrive))
    print fmtinfo.format("Model", str(model))
    print fmtinfo.format("Serial", str(serial))
    print fmtinfo.format("AvailCapacity", str(physDriveGiB) + " GiB")
    print fmtinfo.format("TestedCapacity", str(testcapacity) + " GiB")
    print fmtinfo.format("CPU", str(cpu))
    print fmtinfo.format("Cores", str(cpuCores))
    print fmtinfo.format("Frequency", str(cpuFreqMHz))

    print "\n"
    print fmtstr.format("Test Description", "BW(MB/s)", "IOPS", "Lat(us)")
    print fmtstr.format("-"*maxlen, "-"*8, "-"*9, "-"*8)
    for o in oc:
        if o['desc'] == "":
            # This is a header-printing job, don't thread out
            print "\n" + fmtstr.format("---"+o['name']+"---", "", "", "")
            sys.stdout.flush()
            o['cmdline'](o)
        else:
            # This is a real test job, run it in a thread
            if (sys.stdout.isatty()):
                print fmtstr.format(o['desc'], "Runtime", "00:00:00", "..."),
                print "\r",
            else:
                print descfmt.format(o['desc']),
            sys.stdout.flush()
            starttime = datetime.datetime.now()
            job = threading.Thread(target=JobWrapper, kwargs=(o))
            job.start()
            while job.isAlive():
                now = datetime.datetime.now()
                delta = now - starttime
                dstr = "{0:02}:{1:02}:{2:02}".format(delta.seconds / 3600,
                                                     (delta.seconds%3600)/60,
                                                     delta.seconds % 60)
                if (sys.stdout.isatty()):
                    # Blink runtime to make it obvious stuff is happening
                    if (delta.seconds % 2) != 0:
                        print fmtstr.format(o['desc'], "Runtime", dstr, "..."),
                        print "\r",
                    else:
                        print fmtstr.format(o['desc'], "", dstr, "") + "\r",
                sys.stdout.flush()
                time.sleep(1)
            job.join()
            # Pretty-print with grouping, if possible
            try:
                ret_iops = "{:,}".format(int(ret_iops))
                ret_mbps = "{:0,.2f}".format(float(ret_mbps))
            except:
                pass
            if (sys.stdout.isatty()):
                print fmtstr.format(o['desc'], ret_mbps, ret_iops, ret_lat)
            else:
                print " " + resfmt.format(o['desc'], ret_mbps, ret_iops, ret_lat)
            sys.stdout.flush()
            # On any error abort the test, all future results could be invalid
            if ret_mbps == "ERROR":
                print "ERROR DETECTED, ABORTING TEST RUN."
                sys.exit(2)

def GenerateResultODS():
    """Builds a new ODS spreadsheet w/graphs from generated test CSV files."""

    def GetContentXMLFromODS( odssrc ):
        """Extract content.xml from an ODS file, where the sheet lives."""
        ziparchive = zipfile.ZipFile( odssrc )
        content = ziparchive.read("content.xml")
        return content

    def ReplaceSheetWithCSV_regex(sheetName, csvName, xmltext):
        """Replace a named sheet with the contents of a CSV file."""
        newt  = '<table:table table:name='
        newt += '"' + sheetName + '"' + ' table:style-name="ta1" > '
        newt += '<table:table-column table:style-name="co1" '
        newt += 'table:default-cell-style-name="Default"/>'
        # Insert the rows, one entry at a time
        with open(csvName) as f:
            for line in f:
                line = line.rstrip()
                newt += '<table:table-row table:style-name="ro1">'
                for val in line.split(','):
                    try:
                        cell  = '<table:table-cell office:value-type="float" '
                        cell += 'office:value="' + str(float(val))
                        cell += '" calcext:value-type="float"><text:p>'
                        cell += str(float(val)) + '</text:p></table:table-cell>'
                    except: # It's not a float, so let's call it a string
                        cell  = '<table:table-cell office:value-type="string" '
                        cell += 'calcext:value-type="string"><text:p>'
                        cell += str(val) + '</text:p></table:table-cell>'
                    newt += cell
                newt += '</table:table-row>'
            f.close()
        # Close the tags
        newt += '</table:table>'
        # Replace the XML using lazy string matching
        searchstr  = '<table:table table:name="' + sheetName
        searchstr += '.*?</table:table>'
        return re.sub(searchstr, newt, xmltext)

    def UpdateContentXMLToODS_text( odssrc, odsdest, xmltext ):
        """Replace content.xml in an ODS w/an in-memory copy and write new.

        Replace content.xml in an ODS file with in-memory, modified copy and
        write new ODS. Can't just copy source.zip and replace one file, the
        output ZIP file is not correct in many cases (opens in Excel but fails
        ODF validation and LibreOffice fails to load under Windows)
        """
        if os.path.exists(odsdest):
            os.unlink(odsdest)

        # Windows ZipArchive will not use "Store" even with "no compression"
        # so we need to have a mimetype.zip file encoded below to match spec:
        mimetypezip = """
UEsDBAoAAAAAAOKbNUiFbDmKLgAAAC4AAAAIAAAAbWltZXR5cGVhcHBsaWNhdGlvbi92bmQub2Fz
aXMub3BlbmRvY3VtZW50LnNwcmVhZHNoZWV0UEsBAj8ACgAAAAAA4ps1SIVsOYouAAAALgAAAAgA
JAAAAAAAAACAAAAAAAAAAG1pbWV0eXBlCgAgAAAAAAABABgAAAyCUsVU0QFH/eNMmlTRAUf940ya
VNEBUEsFBgAAAAABAAEAWgAAAFQAAAAAAA==
"""
        zipbytes = base64.b64decode( mimetypezip )
        with open(odsdest, 'wb') as f:
            f.write(zipbytes)

        zasrc = zipfile.ZipFile(odssrc, 'r')
        zadst = zipfile.ZipFile(odsdest, 'a', zipfile.ZIP_DEFLATED)
        for entry in zasrc.namelist():
            if entry == "mimetype":
                continue
            elif entry.endswith('/') or entry.endswith('\\'):
                continue
            elif entry == "content.xml":
                zadst.writestr( "content.xml", xmltext)
            else:
                rdbytes = zasrc.read(entry)
                zadst.writestr(entry, rdbytes)
        zasrc.close()
        zadst.close()

    global odssrc, timeseriescsv, testcsv, physDrive, testcapacity, model
    global serial, uname, fioVerString, odsdest

    xmlsrc = GetContentXMLFromODS( odssrc )
    xmlsrc = ReplaceSheetWithCSV_regex( "Timeseries", timeseriescsv, xmlsrc )
    xmlsrc = ReplaceSheetWithCSV_regex( "Tests", testcsv, xmlsrc )
    # OpenOffice doesn't recalculate these cells on load?!
    xmlsrc = xmlsrc.replace( "_DRIVE", str(physDrive) )
    xmlsrc = xmlsrc.replace( "_TESTCAP", str(testcapacity) )
    xmlsrc = xmlsrc.replace ( "_MODEL", str(model) )
    xmlsrc = xmlsrc.replace( "_SERIAL", str(serial) )
    xmlsrc = xmlsrc.replace( "_OS", str(uname) )
    xmlsrc = xmlsrc.replace( "_FIO", str(fioVerString) )
    UpdateContentXMLToODS_text( odssrc, odsdest, xmlsrc )



fio = ""          # FIO executable
fioVerString = "" # FIO self-reported version
physDrive = ""    # Device path to test
utilization = ""  # Device utilization % 1..100
yes = False       # Skip user verification

cpu = ""         # CPU model
cpuCores = ""    # # of cores (including virtual)
cpuFreqMHz = ""  # "Nominal" speed of CPU
uname = ""       # Kernel name/info

physDriveGiB = ""  # Disk size in GiB (2^n)
physDriveGB = ""   # Disk size in GB (10^n)
physDriveBase = "" # Basename (ex: nvme0n1)
testcapacity = ""  # Total GiB to test
model = ""         # Drive model name
serial = ""        # Drive serial number

ds = ""  # Datestamp to appent to files/directories to uniquify
pwd = "" # $CWD

details = ""       # Test details directory
testcsv = ""       # Intermediate test output CSV file
timeseriescsv = "" # Intermediate iostat output CSV file

odssrc = ""  # Original ODS spreadsheet file
odsdest = "" # Generated results ODS spreadsheet file

oc = [] # The list of tests to run

# These globals are used to return the output results of the test thread
# Required because it's difficult to pass back values from a threading.().
ret_iops = 0 # Last test IOPS
ret_mbps = 0 # Last test MBPs
ret_lat = 0  # Last test in microseconds

ParseArgs()
CheckAdmin()
fio = FindFIO()
CheckFIOVersion()
CollectSystemInfo()
CollectDriveInfo()
VerifyContinue()
SetupFiles()
DefineTests()
RunAllTests()
GenerateResultODS()

print "\nCOMPLETED!\nSpreadsheet file: " + odsdest

