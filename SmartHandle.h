#pragma once

template <class T, void* F> class SmartHandle {
protected:
    T _handle;
public:
    SmartHandle() {}
    virtual ~SmartHandle() {
        (void (*)(T))(_handle);
    }
    T operator = (T handle) {
        _handle = handle;
    }
    T* operator & () {
        return &_handle;
    }
    operator T() {
        return _handle;
    }
};