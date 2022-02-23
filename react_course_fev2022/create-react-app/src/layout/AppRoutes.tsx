import React from 'react';
import { Route, Routes } from "react-router-dom"
import { Component8, Component9 } from '../components';
import { NTFCollectionsListPage } from '../containers/NTFCollectionsListPage';

export const AppRoutes = () => {
    return (
        <Routes>
            <Route path="/component8/:defaultMessage" element={<Component8 message="Default message" />} />
            <Route path="/component9" element={<Component9 message="Default message" />} />
            <Route path="/" element={<NTFCollectionsListPage />} />
        </Routes>
    )
}