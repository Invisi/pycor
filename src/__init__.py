import logging
import os

import config
import excel
import utils


def main():
    log = utils.setup_logger(logging.DEBUG)
    # Iterate over configured folders
    for prof_folder in os.listdir(config.BASE_FOLDER):
        prof_folder = os.path.abspath(prof_folder)
        log.info("Accessing folder: %s", prof_folder)

        # Iterate over the subjects folders.
        for subject_folder in os.listdir(prof_folder):
            current_dir = os.path.join(prof_folder, subject_folder)
            if not os.path.isdir(current_dir) or subject_folder in config.FOLDER_IGNORE:
                log.debug('Skipping %s', subject_folder)
                continue
            log.info("Accessing subfolder: %s", current_dir)

            # Try to find and open corrector
            try:
                e = excel.ExcelCorrector.from_subject_folder(current_dir)
                if not e:
                    log.warning('Ignoring folder %s due to missing corrector file', subject_folder)
                else:
                    e.destroy()  # TODO: Destroy after reading data?
            except excel.CorrectorException:
                # traceback.print_exc()
                pass

            log.info("____________")


if __name__ == '__main__':
    main()
