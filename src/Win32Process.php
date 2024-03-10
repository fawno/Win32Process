<?php
	namespace Fawno\Win32Process;

	use Com;
  use VARIANT;

	class Win32Process {
		public const SW_HIDE            = 0;
		public const SW_NORMAL          = 1;
		public const SW_SHOWMINIMIZED   = 2;
		public const SW_SHOWMAXIMIZED   = 3;
		public const SW_SHOWNOACTIVATE  = 4;
		public const SW_SHOW            = 5;
		public const SW_MINIMIZE        = 6;
		public const SW_SHOWMINNOACTIVE = 7;
		public const SW_SHOWNA          = 8;
		public const SW_RESTORE         = 9;
		public const SW_SHOWDEFAULT     = 10;
		public const SW_FORCEMINIMIZE   = 11;

		public const ST_KERNEL_DRIVER       = 0x00000001;
		public const ST_FILE_SYSTEM_DRIVER  = 0x00000002;
		public const ST_ADAPTER             = 0x00000004;
		public const ST_RECOGNIZER_DRIVER   = 0x00000008;
		public const ST_OWN_PROCESS         = 0x00000010;
		public const ST_SHARE_PROCESS       = 0x00000020;
		public const ST_INTERACTIVE_PROCESS = 0x00000100;

		public const EC_NONE         = 0; // User is not notified.
		public const EC_NOTIFIED     = 1; // User is notified.
		public const EC_RESTART      = 2; // System is restarted with the last-known-good configuration.
		public const EC_RESTART_GOOD = 3; // System attempts to start with a good configuration.

		// http://msdn.microsoft.com/en-us/library/windows/desktop/aa394372%28v=vs.85%29.aspx
		//protected $Win32_Process = [
		public const WIN32_PROCESS = [
			'Name',
			'Caption',
			'CommandLine',
			'CreationClassName',
			'CreationDate',
			'CSCreationClassName',
			'CSName',
			'Description',
			'ExecutablePath',
			'ExecutionState',
			'Handle',
			'HandleCount',
			'InstallDate',
			'KernelModeTime',
			'MaximumWorkingSetSize',
			'MinimumWorkingSetSize',
			'OSCreationClassName',
			'OSName',
			'OtherOperationCount',
			'OtherTransferCount',
			'PageFaults',
			'PageFileUsage',
			'ParentProcessId',
			'PeakPageFileUsage',
			'PeakVirtualSize',
			'PeakWorkingSetSize',
			'Priority',
			'PrivatePageCount',
			'ProcessId',
			'QuotaNonPagedPoolUsage',
			'QuotaPagedPoolUsage',
			'QuotaPeakNonPagedPoolUsage',
			'QuotaPeakPagedPoolUsage',
			'ReadOperationCount',
			'ReadTransferCount',
			'SessionId',
			'Status',
			'TerminationDate',
			'ThreadCount',
			'UserModeTime',
			'VirtualSize',
			'WindowsVersion',
			'WorkingSetSize',
			'WriteOperationCount',
			'WriteTransferCount',
		];

		protected $system = null;
		protected $wbemLocator = null;
		protected $wbemServices = null;
		protected $wbemProcess = null;
		protected $wbemConfig = null;

		public function __construct($system = 'localhost', $user = null, $pass = null) {
			$this->system = $system;
			$this->wbemLocator = new com('WbemScripting.SWbemLocator');
			$this->wbemServices = $this->wbemLocator->ConnectServer($this->system, null, $user, $pass);
			$this->wbemProcess = $this->wbemServices->Get('Win32_Process');
			$this->wbemConfig = $this->wbemServices->Get('Win32_ProcessStartup')->SpawnInstance_();
		}

		public function __destruct() {
		}

		// https://msdn.microsoft.com/en-us/library/windows/desktop/aa389390%28v=vs.85%29.aspx
		// StartMode:
		//   Boot      => Device driver started by the operating system loader. This value is valid only for driver services.
		//   System    => Device driver started by the operating system initialization process. This value is valid only for driver services.
		//   Automatic => Service to be started automatically by the Service Control Manager during system startup.
		//   Manual    => Service to be started by the Service Control Manager when a process calls the StartService method.
		//   Disabled  => Service that can no longer be started.
		public function service_create ($Name, $DisplayName, $PathName, $ServiceType = self::ST_OWN_PROCESS, $ErrorControl = self::EC_NONE, $StartMode = 'Manual', $DesktopInteract = true, $StartName = null, $StartPassword = null, $LoadOrderGroup = null, $LoadOrderGroupDependencies = null, $ServiceDependencies = null) {
			$ServiceInst = $this->wbemServices->Get('Win32_BaseService');
			return $ServiceInst->Create($Name, $DisplayName, $PathName, $ServiceType, $ErrorControl, $StartMode, $DesktopInteract, $StartName, $StartPassword, $LoadOrderGroup, $LoadOrderGroupDependencies, $ServiceDependencies);
		}

		public function service_state ($service) {
			$query = 'SELECT * FROM Win32_Service WHERE Name = \'' . $service . '\'';
			$ServiceSet = $this->wbemServices->ExecQuery($query);
			if ($ServiceSet->Count) {
        foreach ($ServiceSet as $ServiceInst) {
          return $ServiceInst->State;
        }
      }

			return false;
		}

		public function service_status ($service) {
			$query = 'SELECT * FROM Win32_Service WHERE Name = \'' . $service . '\'';
			$ServiceSet = $this->wbemServices->ExecQuery($query);
			if ($ServiceSet->Count) {
        foreach ($ServiceSet as $ServiceInst) {
          return $ServiceInst->Status;
        }
      }

			return false;
		}

		// https://msdn.microsoft.com/en-us/library/windows/desktop/aa389960%28v=vs.85%29.aspx
		public function service_delete ($service) {
			$query = 'SELECT * FROM Win32_Service WHERE Name = \'' . $service . '\'';
			$ServiceSet = $this->wbemServices->ExecQuery($query);
			if ($ServiceSet->Count) {
        foreach ($ServiceSet as $ServiceInst) {
          return $ServiceInst->Delete();
        }
      }

			return false;
		}

		// https://msdn.microsoft.com/en-us/library/windows/desktop/aa393660%28v=vs.85%29.aspx
		public function service_start ($service) {
			$query = 'SELECT * FROM Win32_Service WHERE Name = \'' . $service . '\'';
			$ServiceSet = $this->wbemServices->ExecQuery($query);
			if ($ServiceSet->Count) {
        foreach ($ServiceSet as $ServiceInst) {
          return $ServiceInst->StartService();
        }
      }

			return false;
		}

		// https://msdn.microsoft.com/en-us/library/windows/desktop/aa393673%28v=vs.85%29.aspx
		public function service_stop ($service) {
			$query = 'SELECT * FROM Win32_Service WHERE Name = \'' . $service . '\'';
			$ServiceSet = $this->wbemServices->ExecQuery($query);
			if ($ServiceSet->Count) {
        foreach ($ServiceSet as $ServiceInst) {
          return $ServiceInst->StopService();
        }
      }

			return false;
		}

		public function task_list ($filter = null) {
			foreach ($this->wbemServices->InstancesOf('Win32_Process') as $wbemObject) {
				foreach (self::WIN32_PROCESS as $property) {
          $task[$property] = $wbemObject->$property;
        }
				$tasklist[$task['ProcessId']] = $task;
			}

			if ($filter) {
				$tasklist = array_intersect_key($tasklist, array_intersect(array_map('reset', $tasklist), $filter));
			}

			return $tasklist;
		}

		public function task_kill ($task) {
			$query = 'SELECT * FROM Win32_Process WHERE Name = \'' . $task . '\'';
			if (is_numeric($task)) {
        $query = 'SELECT * FROM Win32_Process WHERE ProcessId = \'' . $task . '\'';
      }

			$ProcessList = $this->wbemServices->ExecQuery($query);
			if ($ProcessList->Count) {
				foreach ($ProcessList as $Process) {
          $Process->Terminate();
        }

				do {
					$ProcessList = $this->wbemServices->ExecQuery($query);
				} while ($ProcessList->Count);
				return true;
			} else {
				return false;
			}
		}

		public function task_wait ($task) {
			$query = 'SELECT * FROM Win32_Process WHERE Name = \'' . $task . '\'';
			if (is_numeric($task)) {
        $query = 'SELECT * FROM Win32_Process WHERE ProcessId = \'' . $task . '\'';
      }

			//$query = 'SELECT * FROM Win32_Process WHERE Name = \'' . $task . '\' or ProcessId = \'' . $task . '\'';
			do {
				usleep(100000);
				$ProcessList = $this->wbemServices->ExecQuery($query);
			} while ($ProcessList->Count);
		}

		public function process_startup_information() {
			return $this->wbemConfig;
		}

		// http://msdn.microsoft.com/en-us/library/aa389388%28v=vs.85%29.aspx
		public function process_create($cmd, $cwd = null, $cfg = null) {
			if (empty($cfg)) {
        $cfg = $this->wbemConfig;
      }

			$pid = new VARIANT(0);

			$result = $this->wbemProcess->Create($cmd, $cwd, $cfg, $pid);

      return (int) $pid;
		}

		public function process_terminate() {

		}
	}
