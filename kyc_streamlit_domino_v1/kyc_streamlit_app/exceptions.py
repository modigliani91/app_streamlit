class KYCAppError(RuntimeError):
    """Message français pouvant être montré à l'utilisateur."""

class ConcurrentRun(KYCAppError):
    pass

class InvalidArchive(KYCAppError):
    pass

class EnvironmentNotReady(KYCAppError):
    pass

class ReportUnavailable(KYCAppError):
    pass
