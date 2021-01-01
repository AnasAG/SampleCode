import logging
import psutil
import os
from win32com.client import Dispatch

# this variable is used for local module only.
# On the project level,
logger = logging.getLogger(__name__)

# list of signals to record during scenario test generation
signals_to_record = []


class ToolOneControl(object):
    """
    This class creates an object for automating (controlling) ToolOne tool from python commands

    Args:
        window_visible:
    """

    def __init__(self, window_visible=True):

        logger.info("Connecting to ToolOne...")
        self._instance = None
        try:
            self._instance = Dispatch("ToolOneNG.Application")
            self._instance.MainWindow.Visible = window_visible
        except Exception:
            logger.exception("Could not connect to ToolOne")
            raise

        self._recorder_index = 0

    def open_project(self, file_path):
        """
        This function opens the project in ToolOne tool
        Args:
            file_path: ToolOne project path

        Returns:

        """
        logger.info("Opening project {}...".format(file_path))
        try:
            if self._instance.ActiveProject is None:
                self._instance.OpenProject(file_path)
            else:
                current_open_project = self._instance.ActiveProject.FullPath
                # check if the same project is already open to avoid opening it again
                if current_open_project != file_path:
                    self._instance.OpenProject(file_path)
        except Exception:
            logger.exception("Could not open project {}".format(file_path))
            raise

    def activate_experiment(self, experiment_name):
        """
        Activate ToolOnes experiment
        Args:
            experiment_name: ToolOne experiment name

        Returns: None

        """
        logger.info("Activating experiment {}...".format(experiment_name))
        try:
            if self._instance.ActiveExperiment is None:
                self._instance.ActiveProject.Experiments[experiment_name].Activate()
            else:
                current_experiment = self._instance.ActiveExperiment.Name
                # check if the same experiment is already activated to avoid activating it again
                if current_experiment != experiment_name:
                    self._instance.ActiveProject.Experiments[experiment_name].Activate()
                else:
                    logger.info("Experiment is already activated {}... ".format(experiment_name))
        except Exception:
            logger.exception("Could not activate experiment {}".format(experiment_name))
            raise

    def ToolOne_version(self):
        """ This function gets the version of the current ToolOne tool
        :return: version
        """
        logger.info("Checking dSPACE ToolOne version...")
        try:
            # get version of the current ToolOne tool
            version = self._instance.Version
        except Exception:
            logger.exception("Could not check ToolOne version")
            raise
        return version

    def current_experiment_name(self):
        """ This function returns the name of the current experiment
        :return: experiment_name
        """
        logger.info("Checking the name of the current active experiment...")
        try:
            # get the name of the current active experiment
            current_active_experiment_name = self._instance.ActiveExperiment.Name
        except Exception:
            logger.exception("Could not get the name of current active experiment")
            raise
        return current_active_experiment_name

    def current_project_name(self):
        """ This function returns the name of the current project
        :return: project_name
        """
        # check the name of current project
        logger.info("Checking the name of the current active project...")
        try:
            # get the name of the current active project
            current_active_project_name = self._instance.ActiveProject.Name
        except Exception:
            logger.exception("Could not get the name of current active project")
            raise
        return current_active_project_name

    def current_application_name(self):
        """ This function returns the name of the current application (loaded?) on the platform
        :return: current_active_application_name
        """
        logger.info("Checking the name of the current loaded application...")
        try:
            pass
            # get the name of the current active application
            current_active_application_name = None  # self._instance...
        except Exception:
            logger.exception("Could not get the name of current loaded application")
            raise
        return current_active_application_name

    def online_calibration_state(self):
        """ Check the current ToolOne state for online calibration
        :return: application_state
        """
        logger.info("Getting the current tool state for online calibration...")
        try:
            # gets the current tool state for calibration
            application_state = self._instance.CalibrationManagement.State
        except Exception:
            logger.exception("Could not get the current tool state for online calibration")
            raise
        return application_state

    def start_online_calibration(self):
        """ This function starts the online calibration of an experiment
        :return: None
        """
        logger.info("Starting online calibration...")
        try:
            if self.online_calibration_state() == 0:
                # start online calibration
                self._instance.CalibrationManagement.StartOnlineCalibration()
        except Exception:
            logger.exception("Could not start online calibration")
            raise

    def stop_online_calibration(self):
        """ This function stops the online calibration of an experiment
        :return: None
        """
        logger.info("Stopping online calibration...")
        try:
            if self.online_calibration_state() == 1:
                # stop online calibration
                self._instance.CalibrationManagement.StopOnlineCalibration()
        except Exception:
            logger.exception("Could not stop online calibration")
            raise

    def is_running_measurement(self):
        """ This function check if the measurement is running
        :return: running_measurement
        """
        logger.info("Checking if the system measurement is running...")
        try:
            # check if measurement for current experiment is running
            running_measurement = self._instance.MeasurementDataManagement.IsMeasuring
        except Exception:
            logger.exception("Could not check system measurement state")
            raise
        return running_measurement

    def start_measuring(self):
        """ This function starts measuring the current experiment
        :return: None
        """
        logger.info("Starting measuring for all devices...")
        try:
            # start measuring
            self._instance.MeasurementDataManagement.Start()
        except Exception:
            logger.exception("Could not start measuring")
            raise

    def stop_measuring(self):
        """ This function stops measuring the current experiment
        :return: None
        """
        logger.info("Stopping measuring for all devices...")
        try:
            # stop measuring
            self._instance.MeasurementDataManagement.Stop()
        except Exception:
            logger.exception("Could not stop measuring")
            raise

    def save_project(self):
        """ This function saves the current project without closing (exiting) it
        :return: None
        """
        logger.info("Saving the project...")
        try:
            # save the current project
            self._instance.ActiveProject.Save()
        except Exception:
            logger.exception("Could not save changes of the project")
            raise

    def close_project(self, save_changes=True):
        """ This function closes the current project, with the option to save modifications or not
        :param save_changes:
        :return: None
        """
        logger.info("Closing the project...")
        try:
            # close the current project with/without saving modifications
            self._instance.ActiveProject.Close(SaveChanges=save_changes)
        except Exception:
            logger.exception("Could not close the project")
            raise

    def restart_ToolOne(self, save_changes=False, window_visible=True):
        """
        Restart ToolOne tool.

        Args:
            save_changes (bool):
            window_visible (bool):

        Returns: None
        """
        logger.info("Restarting ToolOne tool...")
        try:
            # quit ToolOne tool
            self.close_ToolOne(save_changes)
            # start ToolOne tool
            self._instance = Dispatch("ToolOneNG.Application")
            # make the ToolOne GUI visible
            self._instance.MainWindow.Visible = window_visible
        except Exception:
            logger.exception("Could not restart ToolOne")
            raise

    def load_application_from_file(self, applicationFullPath):
        """ This functions loads the application from the file location to the ToolOne platform manager
        :param applicationFullPath:
        :return: None
        """
        logger.info("Loading the application on the Platform...")
        try:
            application_name = os.path.basename(os.path.normpath(applicationFullPath))
            is_application_loaded = self._instance.PlatformManagement.Platforms[0].RealTimeApplications.Contains(
                application_name)

            if is_application_loaded == False:
                # load an application on the Platform
                self._instance.PlatformManagement.Platforms[0].LoadRealtimeApplication(applicationFullPath)
        except Exception:
            logger.exception("Could not load the application from the Platform")
            raise

    def unload_application_from_platform(self):
        """
        This function unloads the current application from the VEOS platform
        Returns: None
        """
        logger.info("Unloading the application from the Platform...")
        try:
            # need to stop online calibration before unloading the experiment to avoid a com-error
            self.stop_online_calibration()
            # Unload the application from the Platform
            self._instance.ActiveExperiment.Platforms[0].RealTimeApplication.Unload()
        except Exception:
            logger.exception("Could not unload the application from the Platform")
            raise

    def start_application_on_platform(self):
        """ This function starts the offline simulation application on the Platform
        :return: None
        """
        logger.info("Starting the offline simulation application on the Platform...")
        try:
            # start the application on the Platform
            active_real_time_applications = self._instance.ActiveExperiment.Platforms[0].RealTimeApplication
            if active_real_time_applications != None:
                active_real_time_applications.Start()
            else:
                logger.info("Currently no active real time application available to start")
        except Exception:
            logger.exception("Could not start the application on the Platform")
            raise

    def stop_application_on_platform(self):
        """
        This function stops the current application on VEOS platform
        Returns: None
        """
        logger.info("Stopping the application currently on the Platform...")
        try:
            # stop the application on the Platform
            active_real_time_applications = self._instance.ActiveExperiment.Platforms[0].RealTimeApplication
            # to avoid error, check if there is a loaded active real time application
            if active_real_time_applications != None:
                active_real_time_applications.Stop()
            else:
                logger.info("Currently no active real time application available to stop")
        except Exception:
            logger.exception("Could not stop the application currently running on the Platform")
            raise

    def pause_application_on_platform(self):
        """
        This function pauses the current application on VEOS platform
        Returns: None
        """
        logger.info("Pausing the application currently on the Platform...")
        try:
            # pause the application on the Platform
            active_real_time_applications = self._instance.ActiveExperiment.Platforms[0].RealTimeApplication
            # to avoid error, check if there is a loaded active real time application
            if active_real_time_applications != None:
                active_real_time_applications.Pause()
            else:
                logger.info("Currently no active real time application available to pause")
        except Exception:
            logger.exception("Could not pause the application currently running on the Platform")
            raise

    def state_application_on_platform(self):
        """This function returns the state of the application loaded on the platform. It gives 0 when no application is
        loaded and 1 when an application is loaded.
        :return: state_application
        """
        logger.info("Getting the state of the application currently on the Platform...")
        try:
            # Getting the state of the application currently on the Platform
            state_application = self._instance.ActiveExperiment.Platforms[0].RealTimeApplication.State
        except Exception:
            logger.exception("Could not get the state of the application currently on the Platform")
            raise
        return state_application

    def restart_application(self, applicationFullPath):
        """
        This function restarts the application on the platform by unloading then reloading the application

        Args:
            applicationFullPath: Applications path

        Returns: None

        """
        logger.info("Restaring the application on the platform...")
        try:
            self.unload_application_from_platform()
            self.load_application_from_file(applicationFullPath)
        except Exception:
            logger.exception("Could not restart the application {} on the platform".format(applicationFullPath))
            raise

    def measurement_recorder(self):
        """ This function returns the recorder object with the specified index or name. Usually it is used as the basic
        for other options in the Recorder Class.
        :return: Recorder Collection object
        """
        logging.info("Getting the recorder collection...")
        try:
            return self._instance.MeasurementDataManagement.Recorders[self._recorder_index]
        except Exception:
            logging.exception("Could not get the recorder collection")
            raise

    def enable_measurement_start_condition(self, state):
        """ This function specifies if the start condition is enabled or disabled to start measuring.
        :param state:
        :return: None
        """
        logging.info("Setting the start condition option to start measuring after an event occurs...")
        try:
            self._instance.MeasurementDataManagement.Recorders[self._recorder_index].StartCondition.Enabled = state
        except Exception:
            logging.exception("Could not set the start condition option")
            raise

    def set_measurement_trigger_rules(self, trigger_rules):
        """ This function sets the trigger rules for starting measuring after they occur.
        :param trigger_rules:
        :return: object
        """
        logging.info("Setting the trigger rules to {}".format(trigger_rules))
        try:
            return self._instance.MeasurementDataManagement.TriggerRules[trigger_rules]
        except Exception:
            logging.exception("Could not set the trigger rules to {}".format(trigger_rules))
            raise

    def link_trigger_rules_with_start_measurement(self, trigger_rule):
        """ This function links the trigger rules object with the start of recording the measurements.
        :param trigger_rule:
        :return: None
        """
        logging.info("Linking the trigger rules with the start of recording the measurements...")
        try:
            self._instance.MeasurementDataManagement.Recorders[
                self._recorder_index].StartCondition.Trigger = trigger_rule
        except Exception:
            logging.exception("Could not link the trigger rules object with the start of recording the measurements")
            raise

    def configure_start_conditions_for_measurement(self, WithTrigger, OverwriteExisting):
        """ This function starts the measurement recording according to the specified parameters.
        :param WithTrigger:
        :param OverwriteExisting:
        :return: None
        """
        logging.info("Starting recording the measurements according to the specified parameters...")
        try:
            self._instance.MeasurementDataManagement.Recorders[self._recorder_index].Start(WithTrigger,
                                                                                           OverwriteExisting)
        except Exception:
            logging.exception("Could not start the recording of the measurements")
            raise

    def stop_recording_measurement(self):
        """ This function stops recording the measurements.
        :return: None
        """
        logging.info("Stopping recording the measurements...")
        try:
            # stop the recording
            self._instance.MeasurementDataManagement.Recorders[self._recorder_index].Stop()
        except Exception:
            logging.exception("Could not stop recording the measurements")
            raise

    def stop_measuring_measurement(self):
        """ This function stops running the measurements.
        :return: None
        """
        logging.info("Stopping measuring...")
        try:
            # stop measuring
            self._instance.MeasurementDataManagement.Stop()
        except Exception:
            logging.exception("Could not stop measuring")
            raise

    def read_signals_from_file(self, signals_file_path):
        # This function reads the signals from a file and stores them in a list
        logger.info("Reading signals from signals file...")
        try:
            global signals_to_record
            index = 0
            with open(signals_file_path, "r") as signals_file:
                for line in signals_file:
                    # avoid adding lines starting with #
                    if line.startswith('#'):
                        continue
                    # avoid adding empty lines
                    if line.startswith('\n'):
                        continue
                        # add signal to signals list
                    signals_to_record.append(line)
                    index += 1
        except Exception:
            logger.exception("Could not read signals from file")

    def set_signals_to_record(self):
        """ This function checks which signals need to be recorded and adds them to the recording list.
        :param : None
        :return: None
        """
        logger.info("Configuring signal recording...")
        try:
            # store the recorder commandline in ToolOne
            recorder = self._instance.MeasurementDataManagement.Recorders[self._recorder_index]
            # store signal configuration commandline in ToolOne
            signal_configuration = self._instance.MeasurementDataManagement.MeasurementConfiguration.Signals
            # for entry in signal_paths:
            for signal in signals_to_record:
                # add signal item to the Signals class with the signal's name
                # signal.strip('\n') is needed for ToolOne to get correct naming of the signals
                new_signal = signal_configuration.Add(signal.strip('\n'))
                # insert the signal name to be recorded
                recorder.Signals.Insert(new_signal)
        except Exception:
            logger.exception("Could not configure signal recording")
            raise

    def start_running_test(self, enable_state, trigger_rules, with_trigger, overwrite_existing):
        """ This function contains many basic functions in a sequence to start running online
        calibration and recording the measurements.
        :return: None
        """
        logger.info("Starting test run...")
        try:
            # stop the online calibration
            self.stop_online_calibration()

            self.start_application_on_platform()

            self.start_online_calibration()

            self.enable_measurement_start_condition(enable_state)

            trigger = self.set_measurement_trigger_rules(trigger_rules)

            self.link_trigger_rules_with_start_measurement(trigger)

            self.configure_start_conditions_for_measurement(with_trigger, overwrite_existing)

        except Exception:
            logger.exception("Could not start test run")
            raise

    def stop_recording_and_measuring(self):
        """ This function stops measuring the current experiment
        :return: None
        """
        logger.info("Stopping test run...")
        try:
            # stop recording
            self.stop_recording_measurement()
            # stop the measurement
            self.stop_measuring_measurement()
            # self._instance.MeasurementDataManagement.Stop()
        except Exception:
            logger.exception("Could not stop test run")
            raise

    def get_recording_path(self):
        """ This function gets the path for the signals which will be recorded during the test
        :return: recording path
        """
        logger.info("Getting recording path...")
        try:
            # return the signals going to be recorded during the test
            return self._instance.MeasurementDataManagement.Recorders[self._recorder_index].LastRecordedFiles[0]
        except Exception:
            logger.exception("Could not get recording path")
            raise

    def close_ToolOne(self, save_changes=False):
        """ This function quits ToolOne tool
        :return: None
        """
        logger.info("Closing ToolOne...")
        try:
            # check if any application is loaded on the platform
            # loaded_applications = self._instance.ActiveExperiment.Platforms[0].LoadableApplications.Count
            applications_state = self.state_application_on_platform()
            # if there is at least one uploaded application on the platform stop it
            if applications_state > 0:
                # stop all running applications on platforms before closing ToolOne tool
                for platform in self._instance.ActiveExperiment.Platforms:
                    # stop each platform
                    platform.RealTimeApplication.Stop()
            # quit ToolOne tool
            self._instance.Quit(save_changes)
        except Exception:
            logger.exception("Could not close ToolOne Normally. Trying to kill the process...")
            for process in psutil.process_iter():
                if process.name() == "ToolOne.exe":
                    process.kill()
            raise


def logger_setup():
    # logger = logging.getLogger(__name__)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(name)s - %(funcName)s - %(message)s")

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    file_handler = logging.FileHandler("example.log")
    file_handler.setFormatter(formatter)

    root_logger = logging.getLogger("")
    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)

    root_logger.setLevel(logging.INFO)


if __name__ == '__main__':
    pass
