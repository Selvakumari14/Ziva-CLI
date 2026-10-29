ARCHS_IOS = x86_64-apple-ios aarch64-apple-ios aarch64-apple-ios-sim
ARCHS_ANDROID = aarch64-linux-android armv7-linux-androideabi x86_64-linux-android i686-linux-android 
LIB = libtimon.a
DYLIB = libtimon.so
XCFRAMEWORK = timon.xcframework

all: ios android

ios: $(XCFRAMEWORK)

android: $(ARCHS_ANDROID)
	# After build is done copy files into the android folder
	mkdir -p ../android/app/src/main/jniLibs
	mkdir -p ../android/app/src/main/jniLibs/x86
	mkdir -p ../android/app/src/main/jniLibs/arm64-v8a
	mkdir -p ../android/app/src/main/jniLibs/armeabi-v7a
	mkdir -p ../android/app/src/main/jniLibs/x86_64

	cp ./target/i686-linux-android/release/$(DYLIB) ../android/app/src/main/jniLibs/x86/$(DYLIB)
	cp ./target/aarch64-linux-android/release/$(DYLIB) ../android/app/src/main/jniLibs/arm64-v8a/$(DYLIB)
	cp ./target/armv7-linux-androideabi/release/$(DYLIB) ../android/app/src/main/jniLibs/armeabi-v7a/$(DYLIB)
	cp ./target/x86_64-linux-android/release/$(DYLIB) ../android/app/src/main/jniLibs/x86_64/$(DYLIB)

.PHONY: all ios android $(ARCHS_IOS) $(ARCHS_ANDROID)

$(ARCHS_IOS): %:
	IPHONEOS_DEPLOYMENT_TARGET=17.5 cargo build --target $@ --release

$(ARCHS_ANDROID): %:
	RUSTFLAGS="-C link-arg=-Wl,-z,max-page-size=16384" \
	cargo ndk --target $@ --platform 21 -- build --release

$(XCFRAMEWORK): $(ARCHS_IOS)
	mkdir -p simulator_fat
	lipo -create target/x86_64-apple-ios/release/$(LIB) target/aarch64-apple-ios-sim/release/$(LIB) -output simulator_fat/$(LIB)
	xcodebuild -create-xcframework \
		-library target/aarch64-apple-ios/release/$(LIB) \
		-headers include \
		-library simulator_fat/$(LIB) \
		-headers include \
		-output $@
	cp -r $@ ../ios/$@

