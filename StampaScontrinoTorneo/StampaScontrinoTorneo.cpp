// ================================================================
//  StampaScontrinoTorneo – Thermal receipt printer driver
//
//  Reads a text file and sends it to an ESC/POS network printer.
//  The text file may contain inline control tags (one per line):
//
//  FONT TAGS (set character size)
//    @@FONT:W,H@@    Dynamic: W = width multiplier 0-7, H = height 0-7
//    @@LARGE@@        Preset: 2x width, 2x height
//    @@MEDIUM@@       Preset: 1x width, 2x height
//    @@SMALL@@        Preset: normal (1x / 1x)
//
//  ALIGNMENT TAGS
//    @@CENTER@@       Center-align subsequent text
//    @@LEFT@@         Left-align subsequent text (default)
//
//  PAPER TAGS
//    ====             Feed paper and perform full cut
// ================================================================

#define _CRT_SECURE_NO_WARNINGS
#define _WINSOCK_DEPRECATED_NO_WARNINGS

#include <iostream>
#include <cstring>
#include <cstdlib>
#include <winsock2.h>

#pragma comment(lib, "ws2_32.lib")

// ================================================================
//  Configuration – defaults (can be overridden by env vars)
//    PRINTER_IP   -> env var "PRINTER_IP"   (default: 192.168.80.121)
//    PRINTER_PORT -> env var "PRINTER_PORT" (default: 9100)
// ================================================================
static const char* DEFAULT_PRINTER_IP   = "192.168.80.121";
static const int   DEFAULT_PRINTER_PORT = 9100;

// ================================================================
//  Font size presets
//  Width/height multiplier: 0 = 1x, 1 = 2x, … 7 = 8x
//  Add new entries here and use the tag in your text file.
// ================================================================
struct FontPreset {
	const char* tag;
	int width;
	int height;
};

static const FontPreset fontPresets[] = {
	{ "@@LARGE@@",  1, 1 },   // 2x width, 2x height
	{ "@@MEDIUM@@", 0, 1 },   // 1x width, 2x height
	{ "@@SMALL@@",  0, 0 },   // normal   (1x / 1x)
};
static const int NUM_PRESETS = sizeof(fontPresets) / sizeof(fontPresets[0]);

// ================================================================
//  ESC/POS helper functions
// ================================================================

// GS ! n – set character size
static void setFontSize(SOCKET sock, int width, int height) {
	char cmd[] = { 0x1D, 0x21, (char)((width << 4) | height) };
	send(sock, cmd, sizeof(cmd), 0);
}

// ESC a n – set alignment (0 = left, 1 = center, 2 = right)
static void setAlignment(SOCKET sock, int align) {
	char cmd[] = { 0x1B, 0x61, (char)align };
	send(sock, cmd, sizeof(cmd), 0);
}

// ESC @ – initialize printer
static void initPrinter(SOCKET sock) {
	const char cmd[] = { 0x1B, 0x40 };
	send(sock, cmd, sizeof(cmd), 0);
}

// ESC t n – set code page (16 = WPC1252 / Windows Latin-1)
static void setCodePage(SOCKET sock, int page) {
	const char cmd[] = { 0x1B, 0x74, (char)page };
	send(sock, cmd, sizeof(cmd), 0);
}

// Feed paper + full cut
static void cutPaper(SOCKET sock) {
	const char feed[] = "\r\n\r\n\r\n\r\n\r\n\r\n\r\n\r\n";
	const char cut[]  = { 0x1D, 0x56, 0x00 };
	send(sock, feed, (int)strlen(feed), 0);
	send(sock, cut, sizeof(cut), 0);
}

// ================================================================
//  Tag processing – returns true if the line was a control tag
// ================================================================
static bool processTag(SOCKET sock, const char* line) {

	// ---- Paper cut ----
	if (strncmp(line, "====", 4) == 0) {
		cutPaper(sock);
		return true;
	}

	// ---- Alignment ----
	if (strncmp(line, "@@CENTER@@", 10) == 0) {
		setAlignment(sock, 1);
		printf("Align -> center\n");
		return true;
	}
	if (strncmp(line, "@@LEFT@@", 8) == 0) {
		setAlignment(sock, 0);
		printf("Align -> left\n");
		return true;
	}

	// ---- Dynamic font: @@FONT:W,H@@ ----
	if (strncmp(line, "@@FONT:", 7) == 0) {
		int w = 0, h = 0;
		if (sscanf(line + 7, "%d,%d", &w, &h) == 2) {
			if (w < 0) w = 0; if (w > 7) w = 7;
			if (h < 0) h = 0; if (h > 7) h = 7;
			setFontSize(sock, w, h);
			printf("Font -> dynamic (w=%d, h=%d)\n", w, h);
		}
		return true;
	}

	// ---- Named font presets ----
	for (int i = 0; i < NUM_PRESETS; i++) {
		if (strncmp(line, fontPresets[i].tag, strlen(fontPresets[i].tag)) == 0) {
			setFontSize(sock, fontPresets[i].width, fontPresets[i].height);
			printf("Font -> %s (w=%d, h=%d)\n",
				   fontPresets[i].tag, fontPresets[i].width, fontPresets[i].height);
			return true;
		}
	}

	return false;  // not a tag – treat as printable text
}

// ================================================================
//  Main
// ================================================================
int main(int argc, char* argv[]) {
	if (argc < 2) {
		std::cerr << "Uso: StampaScontrinoTorneo <file>\n";
		return 1;
	}

	// --- Winsock init ---
	WSADATA wsaData;
	if (WSAStartup(MAKEWORD(2, 2), &wsaData) != 0) {
		std::cerr << "Errore WSAStartup\n";
		return 1;
	}

	SOCKET sock = socket(AF_INET, SOCK_STREAM, 0);
	if (sock == INVALID_SOCKET) {
		std::cerr << "Errore socket\n";
		WSACleanup();
		return 1;
	}

	// --- Resolve printer IP and port from env vars (or use defaults) ---
	const char* printerIp = getenv("PRINTER_IP");
	if (!printerIp || printerIp[0] == '\0')
		printerIp = DEFAULT_PRINTER_IP;

	int printerPort = DEFAULT_PRINTER_PORT;
	const char* portEnv = getenv("PRINTER_PORT");
	if (portEnv && portEnv[0] != '\0') {
		int parsed = atoi(portEnv);
		if (parsed > 0 && parsed <= 65535)
			printerPort = parsed;
	}

	printf("Printer: %s:%d\n", printerIp, printerPort);

	// --- Connect to printer ---
	sockaddr_in server{};
	server.sin_family      = AF_INET;
	server.sin_port        = htons(printerPort);
	server.sin_addr.s_addr = inet_addr(printerIp);

	if (connect(sock, (sockaddr*)&server, sizeof(server)) < 0) {
		std::cerr << "Connessione fallita\n";
		closesocket(sock);
		WSACleanup();
		return 1;
	}

	// --- Printer setup ---
	initPrinter(sock);
	setCodePage(sock, 16);          // WPC1252
	setFontSize(sock, 0, 1);        // default: MEDIUM

	// --- Open input file ---
	FILE* file = fopen(argv[1], "r");
	if (!file) {
		std::cerr << "Errore apertura file: " << argv[1] << std::endl;
		closesocket(sock);
		WSACleanup();
		return 1;
	}

	// --- Process lines ---
	char buffer[512];
	while (fgets(buffer, sizeof(buffer), file)) {
		if (processTag(sock, buffer))
			continue;

		printf("buffer:%s", buffer);
		send(sock, buffer, (int)strlen(buffer), 0);
	}

	cutPaper(sock);

	// --- Cleanup ---
	fclose(file);
	closesocket(sock);
	WSACleanup();
	std::cout << "Messaggio inviato.\n";
	return 0;
}
