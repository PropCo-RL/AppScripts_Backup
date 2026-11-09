/**
 * Datei: 04_PaymentFundingService.gs
 * Verantwortlich für: Zahlungsstatus und Bereitstellung zur Wise-Freigabe
 */

/**
 * Überprüft den Transfer-Status oder bereitet die Freigabe vor
 */
function pruefeUndAktualisiereZahlung(profilId, transferId) {
  const transferStatus = holeWiseTransferStatus(transferId);
  if (!transferStatus) {
    return { erfolg: false, meldung: 'Transfer-Status konnte nicht ermittelt werden.' };
  }

  // Für europäische Business-Konten wartet der Transfer auf manuelle Guthaben-Freigabe
  if (transferStatus.status === 'incoming_payment_waiting') {
    return { 
      erfolg: true, 
      status: 'AUSTEHEND_FREIGABE', 
      meldung: `Bereit zur Freigabe in Wise App/Web (Transfer ID: ${transferId})` 
    };
  }

  if (transferStatus.status === 'processing' || transferStatus.status === 'outgoing_payment_sent') {
    return { 
      erfolg: true, 
      status: 'BEZAHLT', 
      meldung: `Bezahlt (Transfer ID: ${transferId}, Status: ${transferStatus.status})` 
    };
  }

  return { 
    erfolg: true, 
    status: transferStatus.status, 
    meldung: `Transfer ID: ${transferId} (Status: ${transferStatus.status})` 
  };
}

/**
 * Fragt den aktuellen Status eines Transfers bei Wise ab (GET /v1/transfers/{id})
 */
function holeWiseTransferStatus(transferId) {
  const apiKey = holeWiseKonfigurationsWert('wiseKey');
  if (!apiKey || !transferId) return null;

  const basisUrl = holeWiseBasisUrl();
  const url = `${basisUrl}/v1/transfers/${transferId}`;

  const antwort = fuehreWiseApiAufrufAus(url, 'GET', apiKey, null);
  if (!antwort || !antwort.id) {
    return null;
  }

  return antwort;
}