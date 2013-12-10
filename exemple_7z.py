# -*- coding: utf-8 -*-
import os
import sys
import logging
import argparse
import subprocess
from datetime import datetime, date, timedelta

ARCHIVE_FILENAME_PATTERN = "py_backup_*_*.7z"
FULL_ARCHIVE_PREFIX = "py_backup_full_"
DIFF_ARCHIVE_PREFIX = "py_backup_diff_"
MAX_DAYS_BETWEEN_FULL = 30
MAX_DAYS_FOR_DIFF = 30

logger = logging.getLogger(__name__)


def parse_date_in_filename(filename):
    import re
    date_in_name_pattern = '_[0-9]{4}-[0-9]{2}-[0-9]{2}'
    match = re.search(date_in_name_pattern, filename)
    if (match):
        date_object = datetime.strptime(match.group(0), '_%Y-%m-%d')
        return (True, date_object)
    else:
        # log that we could not match this filename
        return (False, None)


def get_valid_dir(parser, arg):
    if not os.path.isdir(arg):
        parser.error("The directory %s could not be found!" % arg)
    else:
        return arg


def archive_ok(file_path):
    """Tests whether the archive can be sucessfully decompressed with 7zip."""
    if (not os.path.isfile(file_path)):
        file_path = file_path + '.7z'

    test_command_string = u'7z.exe t '

    dest_path_string = file_path

    full_command = test_command_string + '\"' + dest_path_string + '\" '
    logger.debug("Testing archive with command: %s", full_command)
    encoded_form = full_command.encode(sys.getfilesystemencoding())

    return_value = subprocess.call(encoded_form, shell=False)

    if (return_value == 0):
        logger.info('Backup successful')
        return True
    else:
        logger.warn('Archive file test failed')
        return False


def run_backup(input_dirs, dest_dir, last_full_name, differential=False):
    """ Runs the actual backup using 7zip in archive or update mode."""
    current_date = date.today()

    input_dir_string = ' '.join("\"%s\"" % x for x in input_dirs)

    if differential:
        filename = DIFF_ARCHIVE_PREFIX  + str(current_date)

        diff_path = os.path.join(dest_dir, filename)
        dest_path = os.path.join(dest_dir, last_full_name)

        try:
            os.remove(diff_path + ".7z")
        except OSError:
            pass

        differential_command_string = u'7z.exe u '
        options = " -ms=off -mx=5 -t7z -u- -up0q3r2x2y2z0w2!"

        full_command = differential_command_string + '\"' + dest_path + '\" ' \
                       + input_dir_string + options + '\"' + diff_path + '\"'
    else:
        filename = FULL_ARCHIVE_PREFIX + str(current_date)

        dest_path = os.path.join(dest_dir, filename)
        archive_command_string = u"7z.exe a "

        full_command = archive_command_string + '\"' + dest_path + "\" " \
                       + input_dir_string

    logger.info("Executing command: %s", full_command)
    encoded_form = full_command.encode(sys.getfilesystemencoding())
    return_value = subprocess.call(encoded_form, shell=False)

    if (return_value == 0):
        logger.info('Backup successful')
    else:
        logger.warn('Something went wrong, return value of 7zip is %d', return_value)
        # 1 if directory to backup cannot be found
        # 2 means output file does already exist

    if (archive_ok(dest_path)):
        return True
    else:
        return False


def search_backup_files(input_dirs, dest_dir):
    """Collects previous full and differential archives in the destination directory
       and sorts according to their creation date (encoded in file name)
    """
    import glob

    archives_path = os.path.join(dest_dir, ARCHIVE_FILENAME_PATTERN)
    archive_files = glob.glob(archives_path)
    logger.debug('Processing archive files: %s', (' '.join(archive_files)))

    full_backups = []
    diff_backups = []

    for path in archive_files:
        (directory, fname) = os.path.split(path)
        success, file_date = parse_date_in_filename(fname)
        if success:
            if (fname.startswith(FULL_ARCHIVE_PREFIX)):
                full_backups.append((file_date, fname))
            elif (fname.startswith(FULL_ARCHIVE_PREFIX)):
                diff_backups.append((file_date, fname))
            else:
                # skip file names that are not part of this set
                pass

    full_backups.sort()
    diff_backups.sort()
    return (full_backups, diff_backups)


def decide_backup(input_dirs, dest_dir, full_backups, diff_backups):
    """ Decides whether a full backup is required, i.e. the no full backup exists
        or last full backup is too old. Otherwise a differential backup is sufficient.
    """
    differential = False
    # Consider a diff backup if a recent full backup is available
    if (len(full_backups) > 0):
        last_full_date = full_backups[-1][0]
        last_full_name = full_backups[-1][1]
        delta = datetime.today() - last_full_date
        if (delta < timedelta(days=MAX_DAYS_BETWEEN_FULL)):
            logger.info("Only differential backup required, time delta is %s", str(delta))
            differential = True

    run_backup(input_dirs, dest_dir, last_full_name, differential)


def remove_old_diffs(dest_dir, full_backups, diff_backups):
    """ Removes differential backup files which are older than MAX_DAYS_FOR_DIFF."""
    for (file_date, filename) in diff_backups:
        delta = datetime.today() - file_date
        if (delta > timedelta(days=MAX_DAYS_FOR_DIFF)):
            logger.info("Removing old diff-file: %s" %str(filename))
            os.remove(os.path.join(dest_dir, filename))


def main(argv=None):
    if argv is None:
        argv = sys.argv

    parser = argparse.ArgumentParser()
    parser.add_argument(u"-d", u"--dest-dir", help="destination directory",
                        required=True, type=lambda x: get_valid_dir(parser, x))
    parser.add_argument(u"-i", u'--input-dirs', help="directories to be backed up",
                        nargs='+', required=True, type=lambda x: get_valid_dir(parser, x))
    parser.add_argument(u"-r", u'--remove',
                        help="removes differential backups older than 30 days",
                        action='store_true')
    args = parser.parse_args()

    handler = logging.FileHandler('py_backup.log', 'a', 'utf-8')
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    handler.setLevel(logging.DEBUG)

    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)

    # file names can contain non-ascii characters but windows does not use utf-8
    # http://stackoverflow.com/questions/12764589/python-unicode-encoding
    dest_dir = args.dest_dir.decode('mbcs')
    input_dirs = [x.decode('mbcs') for x in args.input_dirs]

    (full_backups, diff_backups) = search_backup_files(input_dirs, dest_dir)
    decide_backup(input_dirs, dest_dir, full_backups, diff_backups)

    if (args.remove):
        remove_old_diffs(args.dest_dir, full_backups, diff_backups)

if __name__ == "__main__":
    sys.exit(main())