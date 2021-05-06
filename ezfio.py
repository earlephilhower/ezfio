#!/usr/bin/python3

"""ezfio 1.9
earlephilhower@yahoo.com

------------------------------------------------------------------------
ezfio is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 2 of the License, or
(at your option) any later version.

ezfio is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with ezfio.  If not, see <http://www.gnu.org/licenses/>.
------------------------------------------------------------------------

Usage:   ./ezfio.py -d </dev/node> [-u <100..1>]
Example: ./ezfio.py -d /dev/nvme0n1 -u 100

This script requires root privileges so must be run as "root" or
via "sudo ./ezfio.py"

Please be sure to have FIO installed, or you will be prompted to install
and re-run the script."""

from __future__ import print_function
import argparse
import base64
from collections import OrderedDict
import datetime
import glob
import json
import os
import platform
import pwd
import re
import shutil
import socket
import subprocess
import sys
import tempfile
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
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE)
    out = proc.stdout.read()
    err = proc.stderr.read()
    code = proc.wait()
    return int(code), out.decode('UTF-8'), err.decode('UTF-8')


def CheckAdmin():
    """Check that we have root privileges for disk access, abort if not."""
    if os.geteuid() != 0:
        sys.stderr.write("Root privileges are required for low-level disk ")
        sys.stderr.write("access.\nPlease restart this script as root ")
        sys.stderr.write("(sudo) to continue.\n")
        sys.exit(1)


def FindFIO():
    """Try the path and the CWD for a FIO executable, return path or exit."""
    # Determine if FIO is in path or CWD
    try:
        ret, out, err = Run(["fio", "-v"])
        if ret == 0:
            return "fio"
    except:
        try:
            ret, out, err = Run(['./fio', '-v'])
            if ret == 0:
                return "./fio"
        except:
            sys.stderr.write("FIO is required to run IO tests.\n")
            sys.stderr.write("The latest versions can be found at ")
            sys.stderr.write("https://github.com/axboe/fio.\n")
            sys.exit(1)


def CheckFIOVersion():
    """Check that we have a version of FIO installed that we can use."""
    global fio, fioVerString, fioOutputFormat
    code, out, err = Run([fio, '--version'])
    try:
        fioVerString = out.split('\n')[0].rstrip()
        ver = out.split('\n')[0].rstrip().split('-')[1].split('.')[0]
        if int(ver) < 2:
            sys.stderr.write("ERROR: FIO version " + ver + " unsupported, ")
            sys.stderr.write("version 2.0 or later required.  Exiting.\n")
            sys.exit(2)
    except:
        sys.stderr.write("ERROR: Unable to determine version of fio " +
                          "installed.  Exiting.\n")
        sys.exit(2)
    # Now see if we can make exceedance charts
    # Can't just try --output-format=json+ because the FIO in Ubuntu 16.04
    # repo doesn't understand it and *silently ignores ir*.  Instead, use
    # the help output to see if "json+" exists at all...
    try:
        code, out, err = Run([fio, '--help'])
        if (code == 0) and ("json+" in out):
            fioOutputFormat = "json+"
    except:
        pass


def CheckAIOLimits():
    """Ensure kernel AIO max transactions is large enough to run test."""
    global aioNeeded
    # If anything fails, silently continue.  FIO will give error if it
    # can't run due to the AIO setting later on.
    try:
        code, out, err = Run(['cat', '/proc/sys/fs/aio-max-nr'])
        if code == 0:
            aiomaxnr = int(out.split("\n")[0].rstrip())
            if aiomaxnr < int(aioNeeded):
                sys.stderr.write(
                    "ERROR: The kernel's maximum outstanding async IO" +
                    "setting (aio-max-nr) is too\n")
                sys.stderr.write("       low to complete the test run.  Required value is " + str(
                    aioNeeded) + ", current is " + str(aiomaxnr) + "\n")
                sys.stderr.write(
                    "       To fix this temporarially, please execute the following command:\n")
                sys.stderr.write(
                    "            sudo sysctl -w fs.aio-max-nr=" + str(aioNeeded) + "\n")
                sys.stderr.write("Unable to continue.  Exiting.\n")
                sys.exit(2)
    except:
        pass


def ParseArgs():
    """Parse command line options into globals."""
    global physDrive, physDriveDict, physDriveTxt, utilization, nullio, isFile
    global outputDest, offset, cluster, yes, quickie, verify, fastPrecond
    global readOnly

    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="A tool to easily run FIO to benchmark sustained "
        "performance of NVME\nand other types of SSD.",
        epilog="""
Requirements:\n
* Root access (log in as root, or sudo {prog})
* No filesytems or data on target device
* FIO IO tester (available https://github.com/axboe/fio)
* sdparm to identify the NVME device and serial number

WARNING: All data on the target device will be DESTROYED by this test.""")
    parser.add_argument("--cluster", dest="cluster", action='store_true',
                        help="Run the test on a cluster (--drive in "+
                        "host1:/dev/p1,host2:/dev/ps,...)", required=False)
    parser.add_argument("--verify", dest="verify", action='store_true',
                        help="Have FIO perform data verifications on reads."+
                        " May impact performance", required=False)
    parser.add_argument("--drive", "-d", dest="physDrive",
                        help="Device to test (ex: /dev/nvme0n1)", required=True)
    parser.add_argument("--utilization", "-u", dest="utilization",
                        help="Amount of drive to test (in percent), 1...100",
                        default="100", type=int, required=False)
    parser.add_argument("--offset", "-s", dest="offset",
                        help="offset from start (in percent), 0...99", default="0",
                        type=int, required=False)
    parser.add_argument("--output", "-o", dest="outputDest",
                        help="Location where results should be saved", required=False)
    parser.add_argument("--yes", dest="yes", action='store_true',
                        help="Skip the final warning prompt (for scripted tests)",
                        required=False)
    parser.add_argument("--fast-precondition", dest='fastpre', action='store_true',
                        help="Only do a single sequential write to precondition drive",
                        required=False)
    parser.add_argument("--quickie", dest="quickie", help=argparse.SUPPRESS,
                        action='store_true', required=False)
    parser.add_argument("--file", dest="file", help="Test using a regular file, not a device",
                        action='store_true', required=False)
    parser.add_argument("--nullio", dest="nullio", help=argparse.SUPPRESS,
                        action='store_true', required=False)
    parser.add_argument("--readonly", dest="readonly", help="Only run read-only tests, don't write to device",
                        action='store_true', required=False)
    args = parser.parse_args()

    physDrive = args.physDrive
    physDriveTxt = physDrive
    utilization = args.utilization
    outputDest = args.outputDest
    offset = args.offset
    yes = args.yes
    quickie = args.quickie
    nullio = args.nullio
    verify = args.verify
    fastPrecond = args.fastpre
    cluster = args.cluster
    isFile = args.file
    readOnly = args.readonly

    # For cluster mode, we add a new physDriveList dict and fake physDrive
    if cluster:
        nodes = physDrive.split(",")
        for node in nodes:
            physDriveDict[node.split(":")[0]] = node.split(":")[1]
        physDrive = nodes[0].split(":")[1]

    if (utilization < 1) or (utilization > 100):
        print("ERROR:  Utilization must be between 1...100")
        parser.print_help()
        sys.exit(1)

    if (offset < 0) or (offset > 99) or (offset+utilization > 100):
        print("ERROR:  offset must be between 0...99 while offset + utilization <= 100")
        parser.print_help()
        sys.exit(1)
    # Sanity check that the selected drive is not mounted by parsing mounts
    # This is not guaranteed to catch all as there's just too many different
    # naming conventions out there.  Let's cover simple HDD/SSD/NVME patterns
    pdispart = (re.match('.*p?[1-9][0-9]*$', physDrive) and
                not re.match('.*/nvme[0-9]+n[1-9][0-9]*$', physDrive))
    hit = ""
    with open("/proc/mounts", "r") as f:
        mounts = f.readlines()
    for l in mounts:
        dev = l.split()[0]
        mnt = l.split()[1]
        if dev == physDrive:
            hit = dev + " on " + mnt  # Obvious exact match
        if pdispart:
            chkdev = dev
        else:
            # /dev/sdp# is special case, don't remove the "p"
            if re.match('^/dev/sdp.*$', dev):
                chkdev = re.sub('[1-9][0-9]*$', '', dev)
            else:
                # Need to see if mounted partition is on a raw device being tested
                chkdev = re.sub('p?[1-9][0-9]*$', '', dev)
        if chkdev == physDrive:
            hit = dev + " on " + mnt
    if hit != "":
        print("ERROR:  Mounted volume '" + str(hit) + "' is on same device" +
              "as tested device '" + str(physDrive) + "'.  ABORTING.")
        sys.exit(2)


def grep(inlist, regex):
    """Implement grep in a non-Pythonic way to make it comprehensible to humans"""
    out = []
    for i in inlist:
        if re.search(regex, i):
            out = out + [i]
    return out


def CollectSystemInfo():
    """Collect some OS and CPU information."""
    global cpu, cpuCores, cpuFreqMHz, uname
    uname = " ".join(platform.uname())
    code, cpuinfo, err = Run(['cat', '/proc/cpuinfo'])
    cpuinfo = cpuinfo.split("\n")
    if 'ppc64' in uname:
        # Implement grep and sed in Python...
        cpu = grep(cpuinfo, r'model')[0].split(': ')[1].replace('(R)', '').replace('(TM)', '')
        cpuCores = len(grep(cpuinfo, r'processor'))
        try:
            code, dmidecode, err = Run(['dmidecode', '--type', 'processor'])
            cpuFreqMHz = int(round(float(grep(dmidecode.split("\n"), r'Current Speed')[0].rstrip().lstrip().split(" ")[2])))
        except:
            cpuFreqMHz = int(round(float(grep(cpuinfo, r'clock')[0].split(': ')[1][:-3])))
    else:
        model_names = grep(cpuinfo, r'model name')
        cpu = model_names[0].split(': ')[1].replace('(R)', '').replace('(TM)', '')
        cpuCores = len(model_names)
        try:
            code, dmidecode, err = Run(['dmidecode', '--type', 'processor'])
            cpuFreqMHz = int(round(float(grep(dmidecode.split("\n"), r'Current Speed')[0].rstrip().lstrip().split(" ")[2])))
        except:
            cpuFreqMHz = int(round(float(grep(cpuinfo, r'cpu MHz')[0].split(': ')[1])))


def VerifyContinue():
    """User's last chance to abort the test.  Exit if they don't agree."""
    if not yes:
        print("-" * 75)
        print("WARNING! " * 9)
        print("THIS TEST WILL DESTROY ANY DATA AND FILESYSTEMS ON " + physDrive)
        cont = input("Please type the word \"yes\" and hit return to " +
                         "continue, or anything else to abort.")
        print("-" * 75 + "\n")
        if cont != "yes":
            print("Performance test aborted, drive is untouched.")
            sys.exit(1)


def CollectDriveInfo():
    """Get important device information, exit if not possible."""
    global physDriveGiB, physDriveGB, physDriveBase, testcapacity, testoffset
    global model, serial, physDrive, isFile
    # We absolutely need this information
    pd = physDrive.split(',')[0]
    try:
        if isFile:
            physDriveBase = os.path.basename(pd)
            physDriveBytes = str(os.stat(pd).st_size) + "\n"
        else:
            physDriveBase = os.path.basename(pd)
            code, physDriveBytes, err = Run(['blockdev', '--getsize64', pd])
            if code != 0:
                raise Exception("Can't get drive size for " + pd)
        physDriveBytes = physDriveBytes.split('\n')[0]
        physDriveBytes = int(physDriveBytes)
        physDriveGB = int(physDriveBytes / (1000 * 1000 * 1000))
        physDriveGiB = int(physDriveBytes / (1024 * 1024 * 1024))
        testcapacity = int((physDriveGiB * utilization) / 100)
        testoffset = int((physDriveGiB * offset) / 100)
    except:
        print("ERROR: Can't get '" + pd + "' size. Incorrect device name?")
        sys.exit(1)
    # These are nice to have, but we can run without it
    model = "UNKNOWN"
    serial = "UNKNOWN"
    try:
        nvmeclicmd = ['nvme', 'list', '--output-format=json']
        code, nvmecli, err = Run(nvmeclicmd)
        if code == 0:
            j = json.loads(nvmecli)
            for drive in j['Devices']:
                if drive['DevicePath'] == pd:
                    model = drive['ModelNumber']
                    serial = drive['SerialNumber']
                    return
    except:
        pass  # An error in nvme is not a problem
    try:
        sdparmcmd = ['sdparm', '--page', 'sn', '--inquiry', '--long', pd]
        code, sdparm, err = Run(sdparmcmd)
        lines = sdparm.split("\n")
        if len(lines) == 4:
            model = re.sub(
                r'\s+', " ", lines[0].split(":")[1].lstrip().rstrip())
            serial = re.sub(r'\s+', " ", lines[2].lstrip().rstrip())
        else:
            print("Unable to identify drive using sdparm. Continuing.")
    except:
        print("Install sdparm to allow model/serial extraction. Continuing.")


def CSVInfoHeader(f):
    """Headers to the CSV file (ending up in the ODS at the test end)."""
    global physDriveTxt, model, serial, physDriveGiB, testcapacity, testoffset
    global cpu, cpuCores, cpuFreqMHz, uname, quickie, fastPrecond
    if quickie:
        prefix = "QUICKIE-INVALID-RESULTS-"
    else:
        prefix = ""
    if fastPrecond:
        prefix = "FASTPRECOND-" + prefix
    AppendFile("Drive," + prefix + str(physDriveTxt).replace(",", " "), f)
    AppendFile("Model," + prefix + str(model), f)
    AppendFile("Serial," + prefix + str(serial), f)
    AppendFile("AvailCapacity," + prefix + str(physDriveGiB) + ",GiB", f)
    if offset == 0:
        testcap = str(testcapacity)
    else:
        testcap = str(testcapacity) + " @ " + str(testoffset)
    AppendFile("TestedCapacity," + prefix + str(testcap) + ",GiB", f)
    AppendFile("CPU," + prefix + str(cpu), f)
    AppendFile("Cores," + prefix + str(cpuCores), f)
    AppendFile("Frequency," + prefix + str(cpuFreqMHz), f)
    AppendFile("OS," + prefix + str(uname), f)
    AppendFile("FIOVersion," + prefix + str(fioVerString), f)


def SetupFiles():
    """Set up names for all output/input files, place headers on CSVs."""
    global ds, details, testcsv, timeseriescsv, odssrc, odsdest
    global physDriveBase, fioVerString, outputDest, timeseriesclatcsv
    global timeseriesslatcsv

    # Datestamp for run output files
    ds = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # The unique suffix we generate for all output files
    suffix = str(physDriveGB) + "GB_" + str(cpuCores) + "cores_"
    suffix += str(cpuFreqMHz) + "MHz_" + physDriveBase + "_"
    suffix += socket.gethostname() + "_" + ds

    if not outputDest:
        outputDest = os.getcwd()
    # The "details" directory contains the raw output of each FIO run
    details = outputDest + "/details_" + suffix
    if os.path.exists(details):
        shutil.rmtree(details)
    os.makedirs(details)
    # Copy this script into it for posterity
    shutil.copyfile(__file__, details + "/" + os.path.basename(__file__))

    # Files we're going to generate, encode some system info in the names
    # If the output files already exist, erase them
    testcsv = details + "/ezfio_tests_"+suffix+".csv"
    if os.path.exists(testcsv):
        os.unlink(testcsv)
    CSVInfoHeader(testcsv)
    AppendFile("Type,Write %,Block Size,Threads,Queue Depth/Thread,IOPS," +
               "Bandwidth (MB/s),Read Latency (us),Write Latency (us)," +
               "System CPU,User CPU", testcsv)
    timeseriescsv = details + "/ezfio_timeseries_"+suffix+".csv"
    timeseriesclatcsv = details + "/ezfio_timeseriesclat_"+suffix+".csv"
    timeseriesslatcsv = details + "/ezfio_timeseriesslat_"+suffix+".csv"
    for f in [timeseriescsv, timeseriesclatcsv, timeseriesslatcsv]:
        if os.path.exists(f):
            os.unlink(f)
        CSVInfoHeader(f)
    AppendFile(",".join(["IOPS"] + list(physDriveDict.keys())),
               timeseriescsv)  # Add IOPS header
    hdr = ""
    for host in physDriveDict.keys():
        hdr = hdr + ',' + host + "-read"
        hdr = hdr + ',' + host + "-write"
    AppendFile('CLAT-read,CLAT-write' + hdr,
               timeseriesclatcsv)  # Add IOPS header
    AppendFile('SLAT-read,SLAT-write' + hdr,
               timeseriesslatcsv)  # Add IOPS header

    # ODS input and output files
    odssrc = os.path.dirname(os.path.realpath(__file__)) + "/original.ods"
    if not os.path.exists(odssrc):
        print("ERROR: Can't find original ODS spreadsheet '" + odssrc + "'.")
        sys.exit(1)
    odsdest = outputDest + "/ezfio_results_"+suffix+".ods"
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
        super(FIOError, self).__init__()
        self.cmdline = cmdline
        self.code = code
        self.stderr = stderr
        self.stdout = stdout


def TestName(seqrand, wmix, bs, threads, iodepth):
    """Return full path and filename prefix for test of specified params"""
    global details, physDriveBase
    testfile = str(details) + "/Test" + str(seqrand) + "_w" + str(wmix)
    testfile += "_bs" + str(bs) + "_threads" + str(threads) + "_iodepth"
    testfile += str(iodepth) + "_" + str(physDriveBase) + ".out"
    return testfile


def SequentialConditioning():
    """Sequentially fill the complete capacity of the drive once."""
    global quickie, fastPrecond, nullio, readOnly

    def GenerateJobfile(drive, testcapacity, testoffset):
        """Write the sequential jobfile for a single server"""
        jobfile = tempfile.NamedTemporaryFile(delete=False, mode='w')
        for dr in drive.split(','):
            jobfile.write("[SeqCond-" + dr + "]\n")
            # Note that we can't use regular test runner because this test needs
            # to run for a specified # of bytes, not a specified # of seconds.
            jobfile.write("readwrite=write\n")
            jobfile.write("bs=128k\n")
            if nullio:
                jobfile.write("ioengine=null\n")
            else:
                jobfile.write("ioengine=libaio\n")
            jobfile.write("iodepth=64\n")
            jobfile.write("direct=1\n")
            jobfile.write("filename=" + str(dr) + "\n")
            if quickie:
                jobfile.write("size=1G\n")
            else:
                jobfile.write("size=" + str(testcapacity) + "G\n")
            jobfile.write("thread=1\n")
            jobfile.write("offset=" + str(testoffset) + "G\n")
        jobfile.close()
        return jobfile

    cmdline = [fio]
    if not cluster:
        jobfile = GenerateJobfile(physDrive, testcapacity, testoffset)
        cmdline = cmdline + [jobfile.name]
    else:
        jobfile = []
        for host in physDriveDict.keys():
            newjob = GenerateJobfile(
                physDriveDict[host], testcapacity, testoffset)
            cmdline = cmdline + ['--client=' + str(host), str(newjob.name)]
            jobfile = jobfile + [newjob]
    cmdline = cmdline + ['--output-format=' + str(fioOutputFormat)]

    if not readOnly:
        code, out, err = Run(cmdline)
    else:
        code = 0

    if cluster:
        for job in jobfile:
            os.unlink(job.name)
    else:
        os.unlink(jobfile.name)

    if code != 0:
        raise FIOError(" ".join(cmdline), code, err, out)
    else:
        return "DONE", "DONE", "DONE"


def RandomConditioning():
    """Randomly write entire device for the full capacity"""
    global quickie, nullio, readOnly

    def GenerateJobfile(drive, testcapacity, testoffset):
        """Write the random jobfile"""
        jobfile = tempfile.NamedTemporaryFile(delete=False, mode='w')
        for dr in drive.split(','):
            jobfile.write("[RandCond-" + dr + "]\n")
            # Note that we can't use regular test runner because this test needs
            # to run for a specified # of bytes, not a specified # of seconds.
            jobfile.write("readwrite=randwrite\n")
            jobfile.write("bs=4k\n")
            jobfile.write("invalidate=1\n")
            jobfile.write("end_fsync=0\n")
            jobfile.write("group_reporting=1\n")
            jobfile.write("direct=1\n")
            jobfile.write("filename=" + str(dr) + "\n")
            if quickie:
                jobfile.write("size=1G\n")
            else:
                jobfile.write("size=" + str(testcapacity) + "G\n")
            if nullio:
                jobfile.write("ioengine=null\n")
            else:
                jobfile.write("ioengine=libaio\n")
            jobfile.write("iodepth=256\n")
            jobfile.write("norandommap\n")
            jobfile.write("randrepeat=0\n")
            jobfile.write("thread=1\n")
            jobfile.write("offset=" + str(testoffset) + "G\n")
        jobfile.close()
        return jobfile

    cmdline = [fio]
    if not cluster:
        jobfile = GenerateJobfile(physDrive, testcapacity, testoffset)
        cmdline = cmdline + [jobfile.name]
    else:
        jobfile = []
        for host in physDriveDict.keys():
            newjob = GenerateJobfile(
                physDriveDict[host], testcapacity, testoffset)
            cmdline = cmdline + ['--client=' + str(host), str(newjob.name)]
            jobfile = jobfile + [newjob]
    cmdline = cmdline + ['--output-format=' + str(fioOutputFormat)]

    if not readOnly:
        code, out, err = Run(cmdline)
    else:
        code = 0

    if cluster:
        for job in jobfile:
            os.unlink(job.name)
    else:
        os.unlink(jobfile.name)

    if code != 0:
        raise FIOError(" ".join(cmdline), code, err, out)
    else:
        return "DONE", "DONE", "DONE"


def RunTest(iops_log, seqrand, wmix, bs, threads, iodepth, runtime):
    """Runs the specified test, generates output CSV lines."""
    global cluster, physDriveDict

    # Taken from fio_latency2csv.py - needed to convert funky semi-log to normal latencies
    def plat_idx_to_val(idx, FIO_IO_U_PLAT_BITS=6, FIO_IO_U_PLAT_VAL=64):
        """Convert from lat bucket to real value, for obsolete FIO revisions"""
        # MSB <= (FIO_IO_U_PLAT_BITS-1), cannot be rounded off. Use
        # all bits of the sample as index
        if idx < (FIO_IO_U_PLAT_VAL << 1):
            return idx
        # Find the group and compute the minimum value of that group
        error_bits = (idx >> FIO_IO_U_PLAT_BITS) - 1
        base = 1 << (error_bits + FIO_IO_U_PLAT_BITS)
        # Find its bucket number of the group
        k = idx % FIO_IO_U_PLAT_VAL
        # Return the mean of the range of the bucket
        return base + ((k + 0.5) * (1 << error_bits))

    def WriteExceedance(j, rdwr, outfile):
        """Generate an exceedance CSV for read or write from JSON output."""
        global fioOutputFormat
        if fioOutputFormat == "json":
            return  # This data not present in JSON format, only JSON+
        # Generate a dict of combined bins, either for jobs[0] or client_stats[]
        bins = {}
        ios = 0
        try:
            # Non-cluster case will have jobs, only a single one needed
            ios = j['jobs'][0][rdwr]['total_ios']
            if ('N' in j['jobs'][0][rdwr]['clat_ns']) and (j['jobs'][0][rdwr]['clat_ns']['N'] > 0): 
                bins = j['jobs'][0][rdwr]['clat_ns']['bins']
            else:
                bins = {}
        except:
            # Cluster case will have client_stats to combine
            for client_stats in j['client_stats']:
                if client_stats['jobname'] == 'All clients':
                    # Don't bother looking at combined, bins doesn't exist there
                    continue
                if client_stats[rdwr]['total_ios']:
                    ios = ios + client_stats[rdwr]['total_ios']
                    for k in client_stats[rdwr]['clat_ns']['bins'].keys():
                        try:
                            bins[k] = bins[k] + client_stats[rdwr]['clat_ns']['bins'][k]
                        except:
                            bins[k] = client_stats[rdwr]['clat_ns']['bins'][k]
        #ios = client[rdwr]['total_ios']
        #bins = client[rdwr]['clat_ns']['bins']
        if ios:
            runttl = 0
            # This was changed in 2.99 to be in nanoseconds and to discard the crazy _bits magic
            if float(fioVerString.split('-')[1]) >= 2.99:
                lat_ns = []
                # JSON dict has keys of type string, need a sorted integer list for our work...
                for entry in bins:
                    lat_ns.append(int(entry))
                for entry in sorted(lat_ns):
                    lat_us = float(entry) / 1000.0
                    cnt = int(bins[str(entry)])
                    runttl += cnt
                    pctile = 1.0 - float(runttl) / float(ios)
                    if cnt > 0:
                        AppendFile(
                            ",".join((str(lat_us), str(pctile))), outfile)
            else:
                plat_bits = client[rdwr]['clat']['bins']['FIO_IO_U_PLAT_BITS']
                plat_val = client[rdwr]['clat']['bins']['FIO_IO_U_PLAT_VAL']
                for b in range(0, int(client[rdwr]['clat']['bins']['FIO_IO_U_PLAT_NR'])):
                    cnt = int(client[rdwr]['clat']['bins'][str(b)])
                    runttl += cnt
                    pctile = 1.0 - float(runttl) / float(ios)
                    if cnt > 0:
                        AppendFile(
                            ",".join((str(plat_idx_to_val(b, plat_bits, plat_val)),
                                      str(pctile))), outfile)

    def GenerateJobfile(rw, wmix, bs, drive, testcapacity, runtime, threads, iodepth, testoffset):
        """Make a jobfile for the specified test parameters"""
        global verify, nullio
        jobfile = tempfile.NamedTemporaryFile(delete=False, mode='w')
        for dr in drive.split(","):
            jobfile.write("[test-" + dr + "]\n")
            jobfile.write("readwrite=" + str(rw) + "\n")
            jobfile.write("rwmixwrite=" + str(wmix) + "\n")
            jobfile.write("bs=" + str(bs) + "\n")
            jobfile.write("invalidate=1\n")
            jobfile.write("end_fsync=0\n")
            jobfile.write("group_reporting=1\n")
            jobfile.write("direct=1\n")
            jobfile.write("filename=" + str(dr) + "\n")
            jobfile.write("size=" + str(testcapacity) + "G\n")
            jobfile.write("time_based=1\n")
            jobfile.write("runtime=" + str(runtime) + "\n")
            if nullio:
                jobfile.write("ioengine=null\n")
            else:
                jobfile.write("ioengine=libaio\n")
            jobfile.write("numjobs=" + str(threads) + "\n")
            jobfile.write("iodepth=" + str(iodepth) + "\n")
            jobfile.write("norandommap=1\n")
            jobfile.write("randrepeat=0\n")
            jobfile.write("thread=1\n")
            jobfile.write("exitall=1\n")
            if verify:
                jobfile.write("verify=crc32c\n")
                jobfile.write("random_generator=lfsr\n")
            jobfile.write("offset=" + str(testoffset) + "G\n")
        jobfile.close()
        return jobfile

    def CombineThreadOutputs(suffix, outcsv, lat):
        """Merge all FIO iops/lat logs across all servers"""
        # The lists may be called "iops" but the same works for clat/slat
        iops = [0] * (runtime + extra_runtime)
        # For latencies, need to keep the _w and _r separate
        iops_w = [0] * (runtime + extra_runtime)
        host_iops = OrderedDict()
        host_iops_w = OrderedDict()
        filecnt = 0
        if not cluster:
            pdd = OrderedDict()
            pdd['localhost'] = 1 # Just the single host, faked here
        else:
            pdd = physDriveDict
        for host in pdd.keys():
            host_iops[host] = [0] * (runtime + extra_runtime)
            host_iops_w[host] = [0] * (runtime + extra_runtime)
            if not cluster:
                fileglob = testfile + str(suffix) + '.*log'
            else:
                fileglob = testfile + str(suffix) + '.*.log.' + host
            for filename in glob.glob(fileglob):
                filecnt = filecnt + 1
                catcmdline = ['cat', filename]
                catcode, catout, caterr = Run(catcmdline)
                if catcode != 0:
                    AppendFile("ERROR", testcsv)
                    raise FIOError(" ".join(catcmdline),
                                   catcode, caterr, catout)
                lines = catout.split("\n")
                # Set time 0 IOPS to first values
                riops = 0
                wiops = 0
                nexttime = 0
                for x in range(0, runtime + extra_runtime):
                    if not lat:
                        iops[x] = iops[x] + riops + wiops
                        host_iops[host][x] = host_iops[host][x] + riops + wiops
                    else:
                        iops[x] = iops[x] + riops
                        iops_w[x] = iops_w[x] + wiops
                        host_iops[host][x] = host_iops[host][x] + riops
                        host_iops_w[host][x] = host_iops_w[host][x] + wiops
                    while len(lines) > 1 and (nexttime < x):
                        parts = lines[0].split(",")
                        nexttime = float(parts[0]) / 1000.0
                        if int(lines[0].split(",")[2]) == 1:
                            wiops = int(parts[1])
                        else:
                            riops = int(parts[1])
                        lines = lines[1:]

        # Generate the combined CSV
        with open(outcsv, 'a') as f:
            for cnt in range(int(extra_runtime/2), runtime + extra_runtime):
                if filecnt > 0 and lat:
                    line = str(float(iops[cnt])/float(filecnt))
                    line = line + ',' + str(float(iops_w[cnt])/float(filecnt))
                else:
                    line = str(iops[cnt])
                if len(pdd.keys()) > 1:
                    for host in pdd.keys():
                        if filecnt > 0 and lat:
                            line = line + ',' + \
                                str(float(host_iops[host][cnt])/float(filecnt))
                            line = line + ',' + \
                                str(float(host_iops_w[host]
                                          [cnt])/float(filecnt))
                        else:
                            line = line + "," + str(host_iops[host][cnt])
                f.write(line + "\n")

    # Output file names
    testfile = TestName(seqrand, wmix, bs, threads, iodepth)

    if seqrand == "Seq":
        rw = "rw"
    else:
        rw = "randrw"

    if iops_log:
        extra_runtime = 10
    else:
        extra_runtime = 0

    cmdline = [fio]
    if not cluster:
        jobfile = GenerateJobfile(rw, wmix, bs, physDrive, testcapacity,
                                  runtime + extra_runtime, threads, iodepth, testoffset)
        cmdline = cmdline + [jobfile.name]
        AppendFile("[JOBFILE]", testfile)
        with open(jobfile.name, 'r') as of:
            txt = of.read()
            AppendFile(txt, testfile)
        if iops_log:
            AppendFile("write_iops_log=" + testfile, jobfile.name)
            AppendFile("write_lat_log=" + testfile, jobfile.name)
            AppendFile("log_avg_msec=1000", jobfile.name)
            AppendFile("log_unix_epoch=0", jobfile.name)
    else:
        jobfile = []
        for host in physDriveDict.keys():
            newjob = GenerateJobfile(rw, wmix, bs, physDriveDict[host], testcapacity,
                                     runtime + extra_runtime, threads, iodepth, testoffset)
            cmdline = cmdline + ['--client=' + str(host), str(newjob.name)]
            AppendFile('[JOBFILE-' + str(host) + "]", testfile)
            with open(newjob.name, 'r') as of:
                txt = of.read()
                AppendFile(txt, testfile)
            jobfile = jobfile + [newjob]
            if iops_log:
                AppendFile("write_iops_log=" + testfile, newjob.name)
                AppendFile("write_lat_log=" + testfile, newjob.name)
                AppendFile("log_avg_msec=1000", newjob.name)
                AppendFile("log_unix_epoch=0", newjob.name)

    cmdline = cmdline + ['--output-format=' + str(fioOutputFormat)]

    # There are some NVME drives with 4k physical and logical out there.
    # Check that we can actually do this size IO, OTW return 0 for all
    skiptest = False
    code, out, err = Run(['blockdev', '--getpbsz', str(physDrive.split(',')[0])])
    if code == 0:
        iomin = int(out.split("\n")[0])
        if int(bs) < iomin:
            skiptest = True

    if readOnly and wmix != 0:
        skiptest = True 

    # Silently ignore failure to return min block size, FIO will fail and
    # we'll catch that a little later.
    if skiptest:
        code = 0
        out = "Test not run because block size " + str(bs)
        out += " below iominsize " + str(iomin) + "\n"
        out += "3;" + "0;" * 100 + "\n"  # Bogus 0-filled resulte line
        err = ""
    else:
        code, out, err = Run(cmdline)
    AppendFile("[STDOUT]", testfile)
    AppendFile(out, testfile)
    AppendFile("[STDERR]", testfile)
    AppendFile(err, testfile)

    if cluster:
        for job in jobfile:
            os.unlink(job.name)
    else:
        os.unlink(jobfile.name)

    # Make sure we had successful completion, else note and abort run
    if code != 0:
        AppendFile("ERROR", testcsv)
        raise FIOError(" ".join(cmdline), code, err, out)

    if iops_log:
        CombineThreadOutputs('_iops', timeseriescsv, False)
        CombineThreadOutputs('_clat', timeseriesclatcsv, True)
        CombineThreadOutputs('_slat', timeseriesslatcsv, True)

    rdiops = 0
    wriops = 0
    rlat = 0
    wlat = 0
    syscpu = 0
    usrcpu = 0
    if not skiptest:
        # Chomp anything before the json.
        for i in range(0, len(out)):
            if out[i] == '{':
                out = out[i:]
                break
        j = json.loads(out)

        if cluster and len(physDriveDict.keys()) == 1:
            client = j['client_stats'][0]
        elif cluster:
            for res in j['client_stats']:
                if res['jobname'] == "All clients":
                    client = res
                    break
        else:
            client = j['jobs'][0]

        syscpu = float(client['sys_cpu'])
        usrcpu = float(client['usr_cpu'])

        rdiops = float(client['read']['iops'])
        wriops = float(client['write']['iops'])

        # 'lat' goes to 'lat_ns' in newest FIO JSON formats...ugh
        try:
            rlat = float(client['read']['lat_ns']['mean']) / 1000  # ns->us
        except:
            rlat = float(client['read']['lat']['mean'])
        try:
            wlat = float(client['write']['lat_ns']['mean']) / 1000  # ns->us
        except:
            wlat = float(client['write']['lat']['mean'])

    iops = "{0:0.0f}".format(rdiops + wriops)
    mbps = "{0:0.2f}".format((float((rdiops+wriops) * bs) /
                              (1024.0 * 1024.0)))
    lat = "{0:0.1f}".format(max(rlat, wlat))

    AppendFile(",".join((str(seqrand), str(wmix), str(bs), str(threads),
                         str(iodepth), str(iops), str(mbps), str(rlat),
                         str(wlat), str(syscpu), str(usrcpu))), testcsv)

    if skiptest:
        AppendFile("1,1\n", testfile + ".exc.read.csv")
        AppendFile("1,1\n", testfile + ".exc.write.csv")
    else:
        WriteExceedance(j, 'read', testfile + ".exc.read.csv")
        WriteExceedance(j, 'write', testfile + ".exc.write.csv")

    return iops, mbps, lat


def DefineTests():
    """Generate the work list for the main worker into OC."""
    global oc, quickie, fastPrecond
    # What we're shmoo-ing across
    bslist = (512, 1024, 2048, 4096, 8192, 16384, 32768, 65536, 131072)
    qdlist = (1, 2, 4, 8, 16, 32, 64, 128, 256)
    threadslist = (1, 2, 4, 8, 16, 32, 64, 128, 256)

    shorttime = 120  # Runtime of point tests
    longtime = 1200  # Runtime of long-running tests
    if quickie:
        shorttime = int(shorttime / 10)
        longtime = int(longtime / 10)

    def AddTest(name, seqrand, writepct, blocksize, threads, qdperthread,
                iops_log, runtime, desc, cmdline):
        """Bare usage add a test to the list to execute"""
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
        """Add an individual run to the list of tests to execute"""
        AddTest(testname, seqrand, wmix, bs, threads, iodepth, iops_log,
                runtime, desc, lambda o: {RunTest(o['iops_log'],
                                                  o['seqrand'], o['wmix'],
                                                  o['bs'], o['threads'],
                                                  o['qdperthread'],
                                                  o['runtime'])})

    def AddTestBSShmoo():
        """Add a sequence of tests varying the block size"""
        AddTest(testname, 'Preparation', '', '', '', '', '', '', '',
                lambda o: {AppendFile(o['name'], testcsv)})
        for bs in bslist:
            desc = testname + ", BS=" + str(bs)
            DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, desc,
                      iops_log, runtime)

    def AddTestQDShmoo():
        """Add a sequence of tests varying the queue depth"""
        AddTest(testname, 'Preparation', '', '', '', '', '', '', '',
                lambda o: {AppendFile(o['name'], testcsv)})
        for iodepth in qdlist:
            desc = testname + ", QD=" + str(iodepth)
            DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, desc,
                      iops_log, runtime)

    def AddTestThreadsShmoo():
        """Add a sequence of tests varying the number of threads"""
        AddTest(testname, 'Preparation', '', '', '', '', '', '', '',
                lambda o: {AppendFile(o['name'], testcsv)})
        for threads in threadslist:
            desc = testname + ", Threads=" + str(threads)
            DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, desc,
                      iops_log, runtime)

    AddTest('Sequential Preconditioning', 'Preparation', '', '', '', '', '',
            '', '', lambda o: {})  # Only for display on-screen
    AddTest('Sequential Preconditioning', 'Seq Pass 1', '100', '131072', '1',
            '256', False, '', 'Sequential Preconditioning Pass 1',
            lambda o: {SequentialConditioning()})
    if not fastPrecond:
        AddTest('Sequential Preconditioning', 'Seq Pass 2', '100', '131072', '1',
                '256', False, '', 'Sequential Preconditioning Pass 2',
                lambda o: {SequentialConditioning()})

    testname = "Sustained Multi-Threaded Sequential Read Tests by Block Size"
    seqrand = "Seq"
    wmix = 0
    threads = 1
    runtime = shorttime
    iops_log = False
    iodepth = 256
    AddTestBSShmoo()

    testname = "Sustained Multi-Threaded Random Read Tests by Block Size"
    seqrand = "Rand"
    wmix = 0
    threads = 16
    runtime = shorttime
    iops_log = False
    iodepth = 16
    AddTestBSShmoo()

    testname = "Sequential Write Tests with Queue Depth=1 by Block Size"
    seqrand = "Seq"
    wmix = 100
    threads = 1
    runtime = shorttime
    iops_log = False
    iodepth = 1
    AddTestBSShmoo()

    if not fastPrecond:
        AddTest('Random Preconditioning', 'Preparation', '', '', '', '', '', '',
                '', lambda o: {})  # Only for display on-screen
        AddTest('Random Preconditioning', 'Rand Pass 1', '100', '4096', '1',
                '256', False, '', 'Random Preconditioning',
                lambda o: {RandomConditioning()})
        AddTest('Random Preconditioning', 'Rand Pass 2', '100', '4096', '1',
                '256', False, '', 'Random Preconditioning',
                lambda o: {RandomConditioning()})

    testname = "Sustained 4KB Random Read Tests by Number of Threads"
    seqrand = "Rand"
    wmix = 0
    bs = 4096
    runtime = shorttime
    iops_log = False
    iodepth = 1
    AddTestThreadsShmoo()

    testname = "Sustained 4KB Random mixed 30% Write Tests by Threads"
    seqrand = "Rand"
    wmix = 30
    bs = 4096
    runtime = shorttime
    iops_log = False
    iodepth = 1
    AddTestThreadsShmoo()

    testname = "Sustained Perf Stability Test - 4KB Random 30% Write"
    AddTest(testname, 'Preparation', '', '', '', '', '', '', '',
            lambda o: {AppendFile(o['name'], testcsv)})
    seqrand = "Rand"
    wmix = 30
    bs = 4096
    runtime = longtime
    iops_log = True
    iodepth = 1
    threads = 256
    DoAddTest(testname, seqrand, wmix, bs, threads, iodepth, testname,
              iops_log, runtime)

    testname = "Sustained 4KB Random Write Tests by Number of Threads"
    seqrand = "Rand"
    wmix = 100
    bs = 4096
    runtime = shorttime
    iops_log = False
    iodepth = 1
    AddTestThreadsShmoo()

    testname = "Sustained Multi-Threaded Random Write Tests by Block Size"
    seqrand = "Rand"
    wmix = 100
    runtime = shorttime
    iops_log = False
    iodepth = 16
    threads = 16
    AddTestBSShmoo()


def RunAllTests():
    """Iterate through the OC work queue and run each job, show progress."""
    global ret_iops, ret_mbps, ret_lat, fioVerString

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
            print("\nFIO Error!\n" + e.cmdline + "\nSTDOUT:\n" + e.stdout)
            print("STDERR:\n" + e.stderr)
            raise
        except:
            print("\nUnexpected error while running FIO job.")
            raise

    print("*" * len(fmtstr.format("", "", "", "")))
    print("ezFio test parameters:\n")

    fmtinfo = "{0: >20}: {1}"
    print(fmtinfo.format("Drive", str(physDriveTxt)))
    print(fmtinfo.format("Model", str(model)))
    print(fmtinfo.format("Serial", str(serial)))
    print(fmtinfo.format("AvailCapacity", str(physDriveGiB) + " GiB"))
    print(fmtinfo.format("TestedCapacity", str(testcapacity) + " GiB"))
    print(fmtinfo.format("TestedOffset", str(testoffset) + " GiB"))
    print(fmtinfo.format("CPU", str(cpu)))
    print(fmtinfo.format("Cores", str(cpuCores)))
    print(fmtinfo.format("Frequency", str(cpuFreqMHz)))
    print(fmtinfo.format("FIO Version", str(fioVerString)))

    print("\n")
    print(fmtstr.format("Test Description", "BW(MB/s)", "IOPS", "Lat(us)"))
    print(fmtstr.format("-"*maxlen, "-"*8, "-"*9, "-"*8))
    for o in oc:
        if o['desc'] == "":
            # This is a header-printing job, don't thread out
            print("\n" + fmtstr.format("---"+o['name']+"---", "", "", ""))
            sys.stdout.flush()
            o['cmdline'](o)
        else:
            # This is a real test job, run it in a thread
            if sys.stdout.isatty():
                print(fmtstr.format(o['desc'], "Runtime", "00:00:00", "..."), end='\r')
            else:
                print(descfmt.format(o['desc']), end='')
            sys.stdout.flush()
            starttime = datetime.datetime.now()
            job = threading.Thread(target=JobWrapper, kwargs=(o))
            job.start()
            while job.is_alive():
                now = datetime.datetime.now()
                delta = now - starttime
                dstr = "{0:02}:{1:02}:{2:02}".format(int(delta.seconds / 3600),
                                                     int((delta.seconds % 3600)/60),
                                                     int(delta.seconds % 60))
                if sys.stdout.isatty():
                    # Blink runtime to make it obvious stuff is happening
                    if (delta.seconds % 2) != 0:
                        print(fmtstr.format(o['desc'], "Runtime", dstr, "..."), end='\r')
                    else:
                        print(fmtstr.format(o['desc'], "", dstr, ""), end='\r')
                sys.stdout.flush()
                time.sleep(1)
            job.join()
            # Pretty-print with grouping, if possible
            try:
                ret_iops = "{:,}".format(int(ret_iops))
                ret_mbps = "{:0,.2f}".format(float(ret_mbps))
            except:
                pass
            if sys.stdout.isatty():
                print(fmtstr.format(o['desc'], ret_mbps, ret_iops, ret_lat))
            else:
                print(" " + resfmt.format(o['desc'],
                                          ret_mbps, ret_iops, ret_lat))
            sys.stdout.flush()
            # On any error abort the test, all future results could be invalid
            if ret_mbps == "ERROR":
                print("ERROR DETECTED, ABORTING TEST RUN.")
                sys.exit(2)


def GenerateResultODS():
    """Builds a new ODS spreadsheet w/graphs from generated test CSV files."""

    def GetContentXMLFromODS(odssrc):
        """Extract content.xml from an ODS file, where the sheet lives."""
        ziparchive = zipfile.ZipFile(odssrc)
        content = ziparchive.read("content.xml").decode('UTF-8')
        content = content.replace("\n", "")
        return content

    def CSVtoXMLSheet(sheetName, csvName):
        """Replace a named sheet with the contents of a CSV file."""
        newt = '<table:table table:name='
        newt += '"' + sheetName + '"' + ' table:style-name="ta1" > '
        newt += '<table:table-column table:style-name="co1" '
        newt += 'table:default-cell-style-name="Default"/>'
        # Insert the rows, one entry at a time
        with open(csvName, 'r') as f:
            for line in f:
                line = line.rstrip()
                newt += '<table:table-row table:style-name="ro1">'
                for val in line.split(','):
                    try:
                        cell = '<table:table-cell office:value-type="float" '
                        cell += 'office:value="' + str(float(val))
                        cell += '"><text:p>'
                        cell += str(float(val)) + \
                            '</text:p></table:table-cell>'
                    except:  # It's not a float, so let's call it a string
                        cell = '<table:table-cell office:value-type="string" '
                        cell += '><text:p>'
                        cell += str(val) + '</text:p></table:table-cell>'
                    newt += cell
                newt += '</table:table-row>'
            f.close()
        # Close the tags
        newt += '</table:table>'
        return newt

    def ReplaceSheetWithCSV_regex(sheetName, csvName, xmltext):
        """Replace a named sheet with the contents of a CSV file."""
        newt = CSVtoXMLSheet(sheetName, csvName)

        # Replace the XML using lazy string matching
        searchstr = '<table:table table:name="' + sheetName
        searchstr += '".*?</table:table>'
        return re.sub(searchstr, newt, xmltext, flags=re.DOTALL)

    def AppendSheetFromCSV(sheetName, csvName, xmltext):
        """Add a new sheet to the XML from the CSV file."""
        newt = CSVtoXMLSheet(sheetName, csvName)

        # Replace the XML using lazy string matching
        searchstr = '<table:named-expressions/>'
        return re.sub(searchstr, newt + searchstr, xmltext, flags=re.DOTALL)

    def UpdateContentXMLToODS_text(odssrc, odsdest, xmltext):
        """Replace content.xml in an ODS w/an in-memory copy and write new.

        Replace content.xml in an ODS file with in-memory, modified copy and
        write new ODS. Can't just copy source.zip and replace one file, the
        output ZIP file is not correct in many cases (opens in Excel but fails
        ODF validation and LibreOffice fails to load under Windows).

        Also strips out any binary versions of objects and the thumbnail,
        since they are no longer valid once we've changed the data in the
        sheet.
        """
        if os.path.exists(odsdest):
            os.unlink(odsdest)

        # Windows ZipArchive will not use "Store" even with "no compression"
        # so we need to have a mimetype.zip file encoded below to match spec:
        mimetypezip = """
UEsDBBQAAAgAAICyN0+FbDmKLgAAAC4AAAAIAAAAbWltZXR5cGVhcHBsaWNhdGlvbi92bmQub2Fz
aXMub3BlbmRvY3VtZW50LnNwcmVhZHNoZWV0UEsBAhQAFAAACAAAgLI3T4VsOYouAAAALgAAAAgA
AAAAAAAAAAAAAAAAAAAAAG1pbWV0eXBlUEsFBgAAAAABAAEANgAAAFQAAAAAAA==
"""
        zipbytes = base64.b64decode(mimetypezip)
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
                zadst.writestr("content.xml", xmltext)
            elif ("Object" in entry) and ("content.xml" in entry):
                # Remove <table:table table:name="local-table"> table
                rdbytes = zasrc.read(entry).decode('UTF-8')
                outbytes = re.sub(
                    '<table:table table:name="local-table">.*</table:table>', "", rdbytes, flags=re.DOTALL)
                zadst.writestr(entry, outbytes)
            elif entry == "META-INF/manifest.xml":
                # Remove ObjectReplacements from the list
                rdbytes = zasrc.read(entry).decode('UTF-8')
                outbytes = ""
                lines = rdbytes.split("\n")
                for line in lines:
                    if not (("ObjectReplacement" in line) or ("Thumbnails" in line)):
                        outbytes = outbytes + line + "\n"
                zadst.writestr(entry, outbytes)
            elif ("Thumbnails" in entry) or ("ObjectReplacement" in entry):
                # Skip binary versions
                continue
            else:
                rdbytes = zasrc.read(entry)
                zadst.writestr(entry, rdbytes)
        zasrc.close()
        zadst.close()

    def CombineExceedanceCSV(qdList, testType, testWpct, testBS, testIOdepth, suffix):
        """Merge multiple exceedance CSVs into a single output file.

        Column merge multiple CSV files into a single one.  Complicated by
        the fact that the number of columns in each may vary.
        """
        csv = details + "/ezfio_exceedance_"+suffix+".csv"
        if os.path.exists(csv):
            os.unlink(csv)
        CSVInfoHeader(csv)
        line1 = ""
        line2 = ""
        for qd in qdList:
            line1 = line1 + \
                ("QD%d Read Exceedance,,QD%d Write Exceedance,,," % (qd, qd))
            line2 = line2 + "rdusec,rdpct,wrusec,wrpct,,"
        AppendFile(line1, csv)
        AppendFile(line2, csv)

        files = []
        for qd in qdList:
            try:
                r = open(TestName(testType, testWpct, testBS,
                                  qd, testIOdepth) + ".exc.read.csv")
            except:
                r = None
            try:
                w = open(TestName(testType, testWpct, testBS,
                                  qd, testIOdepth) + ".exc.write.csv")
            except:
                w = None
            files.append([r, w])
        while True:
            all_empty = True
            l = ""
            for fset in files:
                if fset[0] is None:
                    a = ""
                else:
                    a = fset[0].readline().strip()
                if fset[1] is None:
                    b = ""
                else:
                    b = fset[1].readline().strip()
                l += (a + ",", ",,")[not a]
                l += (b + ",", ",,")[not b]
                l += ','
                all_empty = all_empty and (not a) and (not b)
            AppendFile(l, csv)
            if all_empty:
                break
        return csv

    global odssrc, timeseriescsv, testcsv, physDrive, testcapacity, model, testoffset
    global serial, uname, fioVerString, odsdest, timeseriesclatcsv, timeseriesslatcsv

    xmlsrc = GetContentXMLFromODS(odssrc)
    xmlsrc = ReplaceSheetWithCSV_regex("Timeseries", timeseriescsv, xmlsrc)
    xmlsrc = ReplaceSheetWithCSV_regex(
        "TimeseriesCLAT", timeseriesclatcsv, xmlsrc)
    xmlsrc = ReplaceSheetWithCSV_regex(
        "TimeseriesSLAT", timeseriesslatcsv, xmlsrc)
    xmlsrc = ReplaceSheetWithCSV_regex("Tests", testcsv, xmlsrc)
    # Potentially add exceedance data if we have it
    if fioOutputFormat == "json+":
        csv = CombineExceedanceCSV(
            [1, 4, 16, 32], "Rand", 30, 4096, 1, "exceedance30")
        xmlsrc = ReplaceSheetWithCSV_regex("Exceedance", csv, xmlsrc)
    # Remove draw:image references to deleted binary previews
    xmlsrc = re.sub("<draw:image.*?/>", "", xmlsrc, flags=re.DOTALL)
    # OpenOffice doesn't recalculate these cells on load?!
    xmlsrc = xmlsrc.replace("_DRIVE", str(physDrive))
    xmlsrc = xmlsrc.replace("_TESTCAP", str(testcapacity))
    xmlsrc = xmlsrc.replace("_MODEL", str(model))
    xmlsrc = xmlsrc.replace("_SERIAL", str(serial))
    xmlsrc = xmlsrc.replace("_OS", str(uname))
    xmlsrc = xmlsrc.replace("_FIO", str(fioVerString))
    UpdateContentXMLToODS_text(odssrc, odsdest, xmlsrc)


fio = ""          # FIO executable
fioVerString = ""  # FIO self-reported version
fioOutputFormat = "json"  # Can we make exceedance charts using JSON+ output?
cluster = False   # Running multiple jobs in a cluster using fio --server
physDrive = ""    # Device path to test
physDriveTxt = ""  # Unadulterated drive line
physDriveDict = OrderedDict()  # Device path to test
utilization = ""  # Device utilization % 1..100
offset = ""       # Test region offset % 0..99
yes = False       # Skip user verification
quickie = False   # Flag to indicate short runs, only for ezfio debugging!
nullio = False    # Flag to do no IO at all, use nullio instead
fastPrecond = False  # Only do 1x sequential write for preconditioning (no random)
verify = False    # Use built-in FIO data verification
readOnly = False  # Only run read-only tests

cpu = ""         # CPU model
cpuCores = ""    # # of cores (including virtual)
cpuFreqMHz = ""  # "Nominal" speed of CPU
uname = ""       # Kernel name/info

physDriveGiB = ""  # Disk size in GiB (2^n)
physDriveGB = ""   # Disk size in GB (10^n)
physDriveBase = ""  # Basename (ex: nvme0n1)
testcapacity = ""  # Total GiB to test
testoffset = ""    # test region offset in GiB
model = ""         # Drive model name
serial = ""        # Drive serial number

ds = ""  # Datestamp to appent to files/directories to uniquify
pwd = ""  # $CWD

details = ""       # Test details directory
testcsv = ""       # Intermediate test output CSV file
timeseriescsv = ""  # Intermediate iostat output CSV file
timeseriesclatcsv = ""  # Intermediate iostat output CSV file
timeseriesslatcsv = ""  # Intermediate iostat output CSV file
exceedancecsv = ""  # Intermediate exceedance output CSV

odssrc = ""  # Original ODS spreadsheet file
odsdest = ""  # Generated results ODS spreadsheet file

oc = []  # The list of tests to run
aioNeeded = 4096  # Minimum AIO kernel setting to run all tests

# These globals are used to return the output results of the test thread
# Required because it's difficult to pass back values from a threading.().
ret_iops = 0  # Last test IOPS
ret_mbps = 0  # Last test MBPs
ret_lat = 0  # Last test in microseconds

if __name__ == "__main__":
    ParseArgs()
    CheckAdmin()
    fio = FindFIO()
    CheckFIOVersion()
    CheckAIOLimits()
    CollectSystemInfo()
    CollectDriveInfo()
    VerifyContinue()
    SetupFiles()
    DefineTests()
    RunAllTests()
    GenerateResultODS()

    print("\nCOMPLETED!\nSpreadsheet file: " + odsdest)
